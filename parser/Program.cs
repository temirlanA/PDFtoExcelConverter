using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace parser
{
    public static class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Beginning");
            string text = ExtractTextFromPdf("test.pdf");
            var models = cleaningData(text);
            var table = getTable(models);
            string executableLocation = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string xslLocation = System.IO.Path.Combine(executableLocation, "test.xlsx");
            ExportToExcel(table, xslLocation);
            Console.ReadKey();
        }

        public static void ExportToExcel(this DataTable dataTable, string excelFilePath = null)
        {
            try
            {
                if (dataTable == null || dataTable.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                var excelApp = new Excel.Application();
                excelApp.Workbooks.Add();

                // single worksheet
                Excel._Worksheet workSheet = excelApp.ActiveSheet;

                // column headings
                for (var i = 0; i < dataTable.Columns.Count; i++)
                {
                    workSheet.Cells[1, i + 1] = dataTable.Columns[i].ColumnName;
                }

                // rows
                for (var i = 0; i < dataTable.Rows.Count; i++)
                {
                    // to do: format datetime values before printing
                    for (var j = 0; j < dataTable.Columns.Count; j++)
                    {
                        workSheet.Cells[i + 2, j + 1] = dataTable.Rows[i][j];
                    }
                }

                // check file path
                if (!string.IsNullOrEmpty(excelFilePath))
                {
                    try
                    {
                        workSheet.SaveAs(excelFilePath);
                        excelApp.Quit();
                        Console.WriteLine("Excel file saved!");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                                            + ex.Message);
                    }
                }
                else
                { // no file path is given
                    excelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }

        public static DataTable getTable(List<MainModel> models)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            table.Columns.Add("Код ТН ВЭД", typeof(string));
            table.Columns.Add("Наименование позиции", typeof(string));
            table.Columns.Add("Доп. ед. изм.", typeof(string));
            table.Columns.Add("Ставка ввозной таможенной пошлины (в процентах от таможенной стоимости либо в евро, либо в долларах США)", typeof(string));
            models.ForEach(model =>
            {
                table.Rows.Add(model.Id,model.PositionName,model.MesurementUnit,model.CustomDuty);
            });
            return table;
        }


        public static string ExtractTextFromPdf(string path)
        {
            using (PdfReader reader = new PdfReader(path))
            {
                StringBuilder text = new StringBuilder();

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }

                var removedText = RemoveLast(text, "9508.");
                removedText.Replace("Ставка ввозной", string.Empty);
                removedText.Replace("таможенной", string.Empty);
                removedText.Replace("Доп.", string.Empty);
                removedText.Replace("Код (в процентах", string.Empty);
                removedText.Replace(" Наименование позиции ед.", string.Empty);
                removedText.Replace("ТН ВЭД от таможенной", string.Empty);
                removedText.Replace("изм.", string.Empty);
                removedText.Replace("стоимости либо", string.Empty);
                removedText.Replace("в евро, либо", string.Empty);
                removedText.Replace("в долларах США)", string.Empty);
                removedText.Replace("пошлины", string.Empty);
                removedText.Replace("ТН ВЭД от", string.Empty);
                
                return removedText.ToString();
            }
        }

        public static List<MainModel> cleaningData(string text)
        {
            var list = new List<string>(text.ToString().Split('\n'));
            list = list.Select(l => l.Trim()).ToList();
            list.RemoveAll(string.IsNullOrWhiteSpace);
            var firstList = new List<string>();
            var secondList = new List<string>();
            list.ForEach(item =>
            {
                item = (Regex.Replace(item, @"\+?", string.Empty)).Trim();
                var connected = Regex.Replace(item, @"\s+", "");
                var resultString = Regex.Match(connected, @"\d+").Value;
                if (resultString.StartsWith("01"))
                {
                    firstList.Add(resultString);

                    switch (resultString.Length)
                    {
                        case 4:
                            secondList.Add(item.Substring(5));
                            break;
                        case 10:
                            secondList.Add(item.Substring(14));
                            break;
                        case 6:
                            secondList.Add(item.Substring(8));
                            break;
                        case 9:
                            secondList.Add(item.Substring(11));
                            break;
                        case 8:
                            secondList.Add(item.Substring(10));
                            break;
                        default:
                            secondList.Add(item);
                            break;
                    }

                }
                else
                {
                    var last = secondList.Last();
                    secondList.RemoveAt(secondList.LastIndexOf(secondList.Last()));
                    secondList.Add($"{last}\n{item}");
                }
            });

            var models = new List<MainModel>();

            for (var i = 0; i < secondList.Count; i++)
            {
                var inLoop = true;
                foreach(var mesurementUnit in MesurementUnitList.MesurementUnits) 
                {
                    if (secondList[i].Contains(mesurementUnit))
                    {
                        var seconPart = secondList[i].Substring(secondList[i].LastIndexOf(mesurementUnit)).Substring(mesurementUnit.Length);
                        var connectedSeconPart = Regex.Replace(seconPart, @"\s+", "");
                        var customDuty = Regex.Match(seconPart, @"\d+").Value;
                        models.Add(new MainModel
                        {
                            Id = firstList[i],
                            PositionName = !string.IsNullOrWhiteSpace(customDuty) ? secondList[i].Replace(mesurementUnit, string.Empty).Replace(customDuty, string.Empty) : secondList[i].Replace(mesurementUnit, string.Empty),
                            MesurementUnit = mesurementUnit.Trim(),
                            CustomDuty = customDuty
                        });
                        inLoop = false;
                        break;
                    }
                };
                if (inLoop)
                {
                    models.Add(new MainModel
                    {
                        Id = firstList[i],
                        PositionName = secondList[i]
                    });
                }
            }
            return models;
        }

        public static StringBuilder RemoveLast(this StringBuilder sb, string value)
        {
            if (sb.Length < 1) return sb;
            sb.Remove(0, sb.ToString().LastIndexOf(value));
            sb.Remove(sb.ToString().LastIndexOf(value), value.Length);
            return sb;
        }
    }
}
