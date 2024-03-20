using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel; 

namespace OutlookAddIn1
{
 
    public class ExcelModel
    {
        public string jsonFilePath { get; set; } 

        public ExcelModel()
        {
            this.jsonFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Outlook", "selectedFile.json");
        }
         
        public Excel.Workbook OpenWorkbook()
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            try
            {
                excelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                excelApp = new Excel.Application();

            }

            excelApp.Visible = true; 

            if (File.Exists(jsonFilePath))
            {

                string jsonString = File.ReadAllText(jsonFilePath);
                Dictionary<string, string> dataDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);

                if (File.Exists(dataDictionary["FilePath"]))
                {
                    foreach (Excel.Workbook _workbook in excelApp.Workbooks)
                    {
                        if (string.Equals(_workbook.FullName, dataDictionary["FilePath"], StringComparison.OrdinalIgnoreCase))
                        {
                            workbook = _workbook;
                            break;
                        }
                    }
                    if (workbook == null)
                    {
                        workbook = excelApp.Workbooks.Open(dataDictionary["FilePath"]);
                    }

                }
            }

            return workbook;
        }

        public Excel.Range FindQuoteNumber(Excel.Range range, string value )
        {
            foreach (Microsoft.Office.Interop.Excel.Range cell in range)
            {
                if (cell.Value != null && cell.Value.ToString() == value)
                {
                    return cell;
                }
            }

            return null;
        }

        public string ChangeProposalNumber( string location, string proposalNumber, Excel.Range JobNumberrange)
        {
            if( Directory.Exists(location)){
                
                return location.Replace(proposalNumber,JobNumberrange.Value); 
            }

            return null; 
        }

        public void UpdateExcel(Microsoft.Office.Interop.Excel.Workbook workbook, string folderPath, string QuoteNumber)
        {
            try
            {
                int lastRow = workbook.Sheets["SH"].Cells[workbook.Sheets["SH"].Rows.Count, "D"].End[Excel.XlDirection.xlUp].Row;
                workbook.Sheets["SH"].Cells[lastRow + 1, 50].Value = folderPath;
                workbook.Sheets["SH"].Cells[lastRow + 1, 4].Value = QuoteNumber;
                workbook.Save();

            }
            finally
            {
                //if (workbook != null)
                //{
                //    workbook.Close();
                //}
            };
        }


    }
}
