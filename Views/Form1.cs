using System;
using System.Net.Http;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.IO;
using ShakenExport.Models;
using System.Data;
using System.Data.Common;
using System.Configuration;
using Newtonsoft.Json;

namespace ShakenExport
{
    public partial class Menu : Form
    {

        public Menu()
        {
            InitializeComponent();
        }
        private static readonly HttpClient client = new HttpClient();

        /// <summary>
        /// APIからデータを取得する
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void ImportData_Click(object sender, EventArgs e)
        {
            // Existing API details
            string apiToken = ConfigurationManager.AppSettings["apiToken"];
            string baseUrl = ConfigurationManager.AppSettings["baseurl"];
            string appId = ConfigurationManager.AppSettings["appId"];

            // Fetch data from the first API
            System.Data.DataTable dt1 = FetchDataFromAPI(apiToken, baseUrl, appId);

            // New API details
            string apiToken2 = ConfigurationManager.AppSettings["apiToken2"];
            string appId2 = ConfigurationManager.AppSettings["appId2"];

            // Fetch data from the second API
            System.Data.DataTable dt2 = FetchDataFromAPI(apiToken2, baseUrl, appId2);

            // Merge the data from the two DataTables on the '登録番号' column
            System.Data.DataTable dtMerged = MergeDataTablesOnRegisterNumber(dt1, dt2);

            // Display the data from the new DataTable in the DataGridView
            dataGridView1.DataSource = dtMerged;
        }

        private System.Data.DataTable FetchDataFromAPI(string apiToken, string baseUrl, string appId)
        {
            // Build the API URL
            string url = $"{baseUrl}/k/v1/records.json?app={appId}";


            // Set the API token in the request headers
            client.DefaultRequestHeaders.Add("X-Cybozu-API-Token", apiToken);

            try
            {
                // Send the API request
                var response = client.GetAsync(url).Result;

                // Check the response
                if (response.IsSuccessStatusCode)
                {
                    // Get the response content
                    var content = response.Content.ReadAsStringAsync().Result;

                    dataGridView1.RowCount = 0;

                    // Parse the JSON response
                    var jsonObject = JObject.Parse(content);
                    var records = jsonObject["records"].Children().ToList();

                    // Create a DataTable to hold the data
                    System.Data.DataTable dt = new System.Data.DataTable();
                    string temprecordno = "";
                    List<string> temprowitems = new List<string>();

                    foreach (var record in records)
                    {
                        foreach (var property in record.Children<JProperty>())
                        {
                            // Set the header
                            if (!dt.Columns.Contains(property.Name))
                            {
                                dt.Columns.Add(property.Name);
                            }

                            // Add a new row when the record number changes
                            if (property.Name == "レコード番号" && temprecordno != property.Value["value"].ToString() && temprowitems.Count > 0)
                            {
                                dt.Rows.Add(temprowitems.ToArray());
                                temprowitems.Clear();
                                temprowitems.Add(property.Value["value"].ToString());
                            }
                            else
                            {
                                temprowitems.Add(property.Value["value"].ToString());
                            }
                        }
                    }

                    if (temprowitems.Any())
                    {
                        dt.Rows.Add(temprowitems.ToArray());
                    }
                    else
                    {
                        // If the API request is not successful, display the error message
                        MessageBox.Show("The API request was not successful. Please check the API token and the app ID.", "API Request Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return dt;
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occur during the API request
                throw new Exception(ex.Message);
            }


            return null;
        }
        private System.Data.DataTable MergeDataTablesOnRegisterNumber(System.Data.DataTable dt1, System.Data.DataTable dt2)
        {
            // Create a new DataTable to hold the merged data
            System.Data.DataTable dtMerged = dt1.Clone();

            // Add columns from the second table that are not in the first
            // dt2 の各列に対して、dt1 に存在しない場合、dtMerged に追加します。
            foreach (DataColumn column in dt2.Columns)
            {
                if (!dtMerged.Columns.Contains(column.ColumnName))
                {
                    dtMerged.Columns.Add(column.ColumnName);
                }

            }

            // Create a lookup from the second table
            // dt2 の各行に対して、dt2 の登録番号をキーとして、行をグループ化します。
            var lookup = dt2.AsEnumerable().ToLookup(row => row.Field<string>("登録番号"));

            // Merge the data
            ////rows2 の各行に対して、dt2 の各列を調べます。column が dt1 に存在しない場合、rowMergedの該当する列に値をコピー
            foreach (DataRow row1 in dt1.Rows)
            {
                var rows2 = lookup[row1.Field<string>("登録番号")];
                DataRow rowMerged = dtMerged.NewRow();
                rowMerged.ItemArray = row1.ItemArray;
                //rowMerged に dt2 の各列を追加します。
                foreach (DataRow row2 in rows2)
                {
                    foreach (DataColumn column in dt2.Columns)
                    {
                        if (dt1.Columns.Contains(column.ColumnName)) continue;
                        rowMerged[column.ColumnName] = row2[column];
                    }
                    //rowMerged に dt2 の各行を追加します。
                rowMerged.ItemArray = row2.ItemArray;
                }

                dtMerged.Rows.Add(rowMerged);
            }

            return dtMerged;
        }
        private void exportexcel_btn_Click(object sender, EventArgs e)
        {
            // Create a DataTable to hold the data
            System.Data.DataTable dt = (System.Data.DataTable)dataGridView1.DataSource;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            string templatePath = ConfigurationManager.AppSettings["templatePath"];
            // Check if the template file exists
            if (!File.Exists(templatePath))
            {
                // If the file does not exist, open a dialog for the user to select a file
                MessageBox.Show("The template file does not exist. Please select a valid file.", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    templatePath = openFileDialog.FileName;
                }
                else
                {
                    // If the user does not select a file, return from the method
                    return;
                }
            }
            // Open the template workbook
            Workbook wb = excel.Workbooks.Open(templatePath);


            // Print each row to a separate sheet
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                // Copy the template sheet

                wb.Worksheets[1].Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
                Worksheet worksheet = new Worksheet();
                worksheet = wb.Worksheets[wb.Worksheets.Count];
                worksheet.Name = Convert.ToString(wb.Worksheets.Count);

                // Print specific cells with specific data

                if (dt.Columns.Contains("登録番号") && dt.Rows[i]["登録番号"] != null)
                    worksheet.Cells[12, 4] = dt.Rows[i]["登録番号"].ToString(); // D12
                if (dt.Columns.Contains("車種") && dt.Rows[i]["車種"] != null)
                    worksheet.Cells[12, 8] = dt.Rows[i]["車種"].ToString(); // H12
                if (dt.Columns.Contains("車名") && dt.Rows[i]["車名"] != null)
                    worksheet.Cells[12, 12] = dt.Rows[i]["車名"].ToString(); // L12
                if (dt.Columns.Contains("形状") && dt.Rows[i]["形状"] != null)
                    worksheet.Cells[16, 12] = dt.Rows[i]["形状"].ToString(); // L16
                if (dt.Columns.Contains("年式") && dt.Rows[i]["年式"] != null)
                    worksheet.Cells[16, 4] = dt.Rows[i]["年式"].ToString(); // D16
                if (dt.Columns.Contains("型式") && dt.Rows[i]["型式"] != null)
                    worksheet.Cells[16, 8] = dt.Rows[i]["型式"].ToString(); // H16
                if (dt.Columns.Contains("台車番号") && dt.Rows[i]["台車番号"] != null)
                    worksheet.Cells[19, 4] = dt.Rows[i]["台車番号"].ToString(); // D19
                if (dt.Columns.Contains("車両重量") && dt.Rows[i]["車両重量"] != null)
                    worksheet.Cells[22, 4] = dt.Rows[i]["車両重量"].ToString(); // D22
                if (dt.Columns.Contains("車両総重量") && dt.Rows[i]["車両総重量"] != null)
                    worksheet.Cells[22, 8] = dt.Rows[i]["車両総重量"].ToString(); // H22
                if (dt.Columns.Contains("長さ") && dt.Rows[i]["長さ"] != null)
                    worksheet.Cells[26, 4] = dt.Rows[i]["長さ"].ToString(); // D26
                if (dt.Columns.Contains("幅") && dt.Rows[i]["幅"] != null)
                    worksheet.Cells[26, 8] = dt.Rows[i]["幅"].ToString(); // H26
                if (dt.Columns.Contains("高さ") && dt.Rows[i]["高さ"] != null)
                    worksheet.Cells[26, 12] = dt.Rows[i]["高さ"].ToString(); // L26
                if (dt.Columns.Contains("乗車定員") && dt.Rows[i]["乗車定員"] != null)
                    worksheet.Cells[29, 4] = dt.Rows[i]["乗車定員"].ToString(); // D29
                if (dt.Columns.Contains("燃料機構") && dt.Rows[i]["燃料機構"] != null)
                    worksheet.Cells[29, 8] = dt.Rows[i]["燃料機構"].ToString(); // H29
                if (dt.Columns.Contains("軸間距離") && dt.Rows[i]["軸間距離"] != null)
                    worksheet.Cells[29, 12] = dt.Rows[i]["軸間距離"].ToString(); // L29
                if (dt.Columns.Contains("総排気量または定格出力") && dt.Rows[i]["総排気量または定格出力"] != null)
                    worksheet.Cells[34, 4] = dt.Rows[i]["総排気量または定格出力"].ToString(); // D32
                if (dt.Columns.Contains("所有者の住所") && dt.Rows[i]["所有者の住所"] != null)
                    worksheet.Cells[39, 7] = dt.Rows[i]["所有者の住所"].ToString(); // G39
                if (dt.Columns.Contains("所有者の氏名") && dt.Rows[i]["所有者の氏名"] != null)
                    worksheet.Cells[40, 7] = dt.Rows[i]["所有者の氏名"].ToString(); // G40
                if (dt.Columns.Contains("使用者の氏名") && dt.Rows[i]["使用者の氏名"] != null)
                    worksheet.Cells[43, 7] = dt.Rows[i]["使用者の氏名"].ToString(); // G43
                if (dt.Columns.Contains("使用者の住所") && dt.Rows[i]["使用者の住所"] != null)
                    worksheet.Cells[42, 7] = dt.Rows[i]["使用者の住所"].ToString(); // G42

            }
            // Specify the path where you want to save the file
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveFileDialog.FilterIndex = 2;
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                wb.SaveAs(saveFileDialog.FileName);
                wb.Close();
                excel.Quit();
                MessageBox.Show("印刷が完了しました！");
            }
        }
    }
}
