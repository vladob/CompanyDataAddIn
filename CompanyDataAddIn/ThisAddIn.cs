using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using System.Windows;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System.IO;
using System.Reflection;
using NLog;


namespace CompanyDataAddIn
{

    public partial class ThisAddIn
    {
        private static readonly HttpClient httpClient = new HttpClient();
        private CompanyDataRibbon ribbon;
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static ClassificationManager classificationManager;
        private List<String> financialStatements = new List<string>();
        private Config config;



        private Config LoadConfig()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "CompanyDataAddIn.config.json";

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                string json = reader.ReadToEnd();
                return JsonConvert.DeserializeObject<Config>(json);
            }
        }

        protected override IRibbonExtension[] CreateRibbonObjects()
        {
            ribbon = new CompanyDataRibbon();
            return new IRibbonExtension[] { ribbon };
        }

        public async void FetchAndInsertData(string cin)
        {
            classificationManager = new ClassificationManager();
            try
            {
                await classificationManager.InitializeAllClassificationsAsync();
                // Now the classifications are loaded into classificationManager._classifications
            }
            catch (Exception ex)
            {
                // Handle exceptions as appropriate, e.g., logging or notifying the user.
                System.Windows.Forms.MessageBox.Show($"Error initializing classifications: {ex.Message}");
            }
            try
            {
                logger.Info($"Fetching data for CIN: {cin}");
                // Step 1: Get the accounting entity ID
                string entityId = await GetAccountingEntityId(cin);
                if (string.IsNullOrEmpty(entityId))
                {
                    MessageBox.Show("No accounting entity found for the given CIN.");
                    return;
                }

                // Step 2: Get the accounting entity data
                JObject entityData = await GetAccountingEntityData(entityId);
                if (entityData == null)
                {
                    MessageBox.Show("Failed to fetch accounting entity data.");
                    return;
                }

                // Step 3: Insert data into Excel
                InsertDataIntoExcel(entityData);


                // Step 4: Get financial statements
                JObject financialStetementData = await GetFinancialStatementData(financialStatements[0]);
                if (financialStetementData == null)
                {
                    MessageBox.Show("Failed to fetch financial statement data, for id: ");
                    return;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private async Task<string> GetAccountingEntityId(string cin)
        {
            string url = $"https://www.registeruz.sk/cruz-public/api/uctovne-jednotky?zmenene-od=2000-01-01&ico={cin}";
            string response = await httpClient.GetStringAsync(url);
            JObject json = JObject.Parse(response);
            return json["id"]?[0]?.ToString();
        }

        private async Task<JObject> GetAccountingEntityData(string entityId)
        {
            string url = $"https://www.registeruz.sk/cruz-public/api/uctovna-jednotka?id={entityId}";
            string response = await httpClient.GetStringAsync(url);
            return JObject.Parse(response);
        }

        private async Task<JObject> GetFinancialStatementData(string statementId)
        {
            string url = $"https://www.registeruz.sk/cruz-public/api/uctovna-zavierka?id={statementId}";
            string response = await httpClient.GetStringAsync(url);
            return JObject.Parse(response);
        }

        private void InsertDataIntoExcel(JObject data)
        {
//            var config = LoadConfig();
            Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Application excelApp = Globals.ThisAddIn.Application;

            excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            excelApp.ScreenUpdating = false;

            // Set column widths
            worksheet.Columns["A"].ColumnWidth = 25;
            worksheet.Columns["B"].ColumnWidth = 30;

            // Insert title row
            Excel.Range rangeA1B1 = worksheet.Range["A1:B1"];
            rangeA1B1.Merge();
            rangeA1B1.Interior.Color = System.Drawing.ColorTranslator.FromHtml(config.TitleRow.BackgroundColor);
            rangeA1B1.Font.Color = System.Drawing.ColorTranslator.FromHtml(config.TitleRow.Font.Color);
            rangeA1B1.Font.Bold = config.TitleRow.Font.Bold;
            rangeA1B1.Font.Size = config.TitleRow.Font.Size;
            rangeA1B1.Value = config.TitleRow.LabelSk; // Use config.TitleRow.LabelEn for English
            rangeA1B1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Insert data
            int row = 2;
            foreach (var item in config.DataConfig.OrderBy(c => c.Order))
            {
                if (data.ContainsKey(item.ApiKey))
                {
                    // Insert label (column A)
                    worksheet.Cells[row, 1] = item.LabelSk; // Use item.LabelEn for English

                    // Insert value (column B)
                    if (item.LookupDictionary != null)
                    {
                        string lookupValue = classificationManager.Lookups[item.LookupDictionary][data[item.ApiKey].ToString()];
                        worksheet.Cells[row, 2] = data[item.ApiKey].ToString() + " - " + lookupValue;
                        // worksheet.Cells[row, 2] = lookupValue;
                    } else
                    {
                        worksheet.Cells[row, 2] = data[item.ApiKey].ToString();
                    }

                    // Apply formatting
                    Excel.Range cell = worksheet.Cells[row, 2];
                    cell.Font.Name = item.Font.Name;
                    cell.Font.Size = item.Font.Size;
                    cell.Font.Bold = item.Font.Bold;
                    cell.Font.Italic = item.Font.Italic;
                    cell.Font.Color = System.Drawing.ColorTranslator.FromHtml(item.Font.Color);
                    cell.Interior.Color = System.Drawing.ColorTranslator.FromHtml(item.BackgroundColor);



                    row++;
                }
            }
            if (data.ContainsKey("idUctovnychZavierok"))
            {
                var financialStatementIds = data["idUctovnychZavierok"].ToList();
                financialStatements.Clear();

                foreach (var fsItem in financialStatementIds)
                {
                    financialStatements.Add(fsItem.ToString());
                }

            }

            // Format data cells
            Excel.Range rangeA2A19 = worksheet.Range["A2:A19"];
            rangeA2A19.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#CAEDFB");

            excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            excelApp.ScreenUpdating = true;


        }
        private async void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Enable TLS 1.2
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            config = LoadConfig();

 //           AddCustomRibbonTab();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
