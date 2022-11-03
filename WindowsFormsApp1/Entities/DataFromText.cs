using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Entities
{
    public class DataFromText
    {
        public IList<Header> Header { get; set; }
        public IList<RowData> RowData { get; set; }

        /// <summary>
        /// Export data DataFromText model
        /// </summary>
        /// <param name="dataCollection"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool ExportToExcel(string filePath)
        {
            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            /// Header  
            if (this.Header.Any())
            {
                workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Row(1).Style.Font.Bold = true;
                for (int i = 0; i < this.Header.Count; i++)
                {
                    string[] rowcolNumber = this.Header[i].Position.Split(',');
                    int rowNumber = int.Parse(rowcolNumber[0]);
                    int colNumber = int.Parse(rowcolNumber[1]);

                    workSheet.Cells[rowNumber, colNumber].Value = this.Header[i].Title;
                }
            }
            // Values
            if (this.RowData.Any())
            {
                for (int i = 0; i < this.RowData.Count; i++)
                {
                    for (int j = 0; j < this.RowData[i].Row.Count; j++)
                    {
                        string[] rowcolNumber = this.RowData[i].Row[j].Position.Split(',');
                        int rowNumber = int.Parse(rowcolNumber[0]);
                        int colNumber = int.Parse(rowcolNumber[1]);
                        workSheet.Cells[rowNumber, colNumber].Value = this.RowData[i].Row[j].Value;
                    }
                    
                }
            }

            if (File.Exists(filePath))
                File.Delete(filePath);
            //Create excel file on physical disk    
            FileStream objFileStrm = File.Create(filePath);
            objFileStrm.Close();

            File.WriteAllBytes(filePath, excel.GetAsByteArray());
            return true;
        }
        /// <summary>
        /// Read data with dataFromTextModel
        /// </summary>
        /// <param name="startRowData"></param>
        /// <param name="startRowHeader"></param>
        /// <param name="filePath"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        public void ReadDataFromText(int startRowData = 1, int startRowHeader = 0, string filePath = "", char separator = ',')
        {
            var fileExtension = new[] { ".txt" };

            FileInfo fileInfo = new FileInfo(filePath);
            if (fileExtension.Contains(fileInfo.Extension))
            {
                using (var reader = new StreamReader(filePath))
                {
                    List<Header> headers = new List<Header>();
                    List<RowData> rowDatas = new List<RowData>();
                    var lineNumber = 1;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        if (!string.IsNullOrEmpty(line))
                        {
                            var values = line.Split(separator);
                            if (startRowHeader > 0 && lineNumber == startRowHeader) // Title Header 
                            {
                                for (int i = 0; i < values.Length; i++)
                                {
                                    var header = new Header
                                    {
                                        Title = values[i].Trim(),
                                        Position = $"{lineNumber},{i+1}"
                                    };
                                    headers.Add(header);
                                }
                            }
                            else if (lineNumber >= startRowData && startRowData > 0) // Value Cell 
                            {
                                var row = new RowData();
                                var cells = new List<CellData>();
                                for (int i = 0; i < values.Length; i++)
                                {
                                    cells.Add(new CellData
                                    {
                                        Value = values[i].Trim(),
                                        Position = $"{lineNumber},{i+1}"
                                    });
                                }
                                row.Row = cells;
                                rowDatas.Add(row);
                            }
                        }
                        lineNumber++;
                    }
                    this.Header = headers;
                    this.RowData = rowDatas;
                    reader.Close();
                }
            }
        }
    }

    public class RowData
    {
        public IList<CellData> Row { get; set; }
    }

    public class Header
    {
        public string Title { get; set; }
        public string Position { get; set; }
    }

    public class CellData
    {
        public string Value { get; set; }
        public string Position { get; set; }
    }



}
