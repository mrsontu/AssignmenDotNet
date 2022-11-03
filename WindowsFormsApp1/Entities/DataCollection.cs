using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace WindowsFormsApp1.Entities
{
    public class DataCollection
    {

        public IList<Data> DataList { get; set; }
        public int Conut { get { return DataList.Count; } }
        public DataCollection()
        {
            this.DataList = new List<Data>();
        }
        public IList<Data> Add(Data data)
        {
            this.DataList.Add(data);
            return this.DataList;
        }
        public IList<Data> Remove(Data data)
        {
            this.DataList.Remove(data);
            return this.DataList;
        }

        /// <summary>
        /// Get Data From Txt
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        public DataCollection GetDataFromTxt(string filePath = "", char separator = ',')
        {
            DataCollection dataCollection = new DataCollection()
            {
                DataList = new List<Data>(),
            };

            if (string.IsNullOrEmpty(filePath))
            {
                filePath = Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Files", "GenerateData.txt");
            }
            FileInfo fileInfo = new FileInfo(filePath);
            if (fileInfo.Exists)
            {
                using (var reader = new StreamReader(filePath))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        if (line != null)
                        {
                            var values = line.Split(separator);
                            var data = new Data();
                            data.Values = new List<string>();
                            data.Component = values[0];
                            data.Parameter = values[1];
                            for (var col = 2; col < values.Length; col++)
                            {
                                data.Values.Add(values[col]);
                            }
                            dataCollection.Add(data);
                        }
                    }
                    reader.Close();
                }
            }
            return dataCollection;
        }


        /// <summary>
        /// Export data collection to excel
        /// </summary>
        /// <param name="dataCollection"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool ExportToExcel(DataCollection dataCollection, string filePath, int startRow = 1)
        {
            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            int rowParameter = startRow + 1;
            int startRowValue = rowParameter + 1;

            if (dataCollection.Conut > 0)
            {
                List<Data> datas = (List<Data>)dataCollection.DataList;
                var dataDistincCoponent = datas.Select(x => x.Component).Distinct().ToList();
                int col = 1;
                for (int i = 0; i < dataDistincCoponent.Count(); i++)
                {
                    /// Component
                    var dataWithComponent = datas.Where(x => x.Component.Equals(dataDistincCoponent[i])).ToList();
                    int quantityParameter = dataWithComponent.Count();
                    workSheet.Cells[startRow, col].Value = dataDistincCoponent[i];
                    if (quantityParameter > 1)
                    {
                        workSheet.Cells[startRow, col, startRow, col + quantityParameter - 1].Merge = true;
                    }

                    workSheet.Cells[startRow, col, startRow, col + quantityParameter - 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    workSheet.Cells[startRow, col, startRow, col + quantityParameter - 1].Style.Fill.BackgroundColor.SetColor(Color.Gray);
                    //Paramters
                    for (int j = 0; j < dataWithComponent.Count(); j++)
                    {
                        int colParameter = col + j;
                        workSheet.Cells[rowParameter, colParameter].Value = dataWithComponent[j].Parameter;
                        workSheet.Cells[rowParameter, colParameter].Style.Font.Bold = true;
                        workSheet.Cells[rowParameter, colParameter].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        workSheet.Cells[rowParameter, colParameter].Style.Fill.BackgroundColor.SetColor(Color.Aqua);
                        workSheet.Cells[rowParameter, colParameter].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        /// Values 
                        for (int indexValue = 0; indexValue < dataWithComponent[j].Values.Count(); indexValue++)
                        {
                            workSheet.Cells[startRowValue + indexValue, colParameter].Value = dataWithComponent[j].Values[indexValue];
                            workSheet.Cells[startRowValue + indexValue, colParameter].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                    }
                    col = quantityParameter + col;
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
        /// Get data collection from excel
        /// </summary>
        /// <param name="dataCollection"></param>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public DataCollection ReadDataFromExcel(string filePath, int startRowHeader = 0)
        {
            DataCollection dataCollection;
            FileInfo fileInfo = new FileInfo(filePath);
            if (fileInfo.Exists)
            {

                int rowParameter = startRowHeader > 0 ? startRowHeader + 1 : startRowHeader;
                int startRowData = rowParameter + 1;

                ExcelPackage package = new ExcelPackage(fileInfo);
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                int rows = worksheet.Dimension.End.Row;
                int cols = worksheet.Dimension.End.Column;
                /// Component 
                string addressCellMerged = string.Empty;
                string componentMerged = string.Empty;
                dataCollection = new DataCollection();

                for (int i = 1; i <= cols; i++)
                {
                    Data data = new Data();
                    int countNotMerged = worksheet.Cells.Where(x => !x.Merge).Count();
                    if (worksheet.Cells[1, 1, 1, cols].Merge == false)
                    {
                        if (startRowHeader == 0)
                        {
                            data.Component = string.Empty;
                            data.Parameter = string.Empty;
                        }
                        else
                        {
                            data.Component = worksheet.Cells[startRowHeader, i].Value?.ToString() ?? string.Empty;
                            data.Parameter = string.Empty;
                            startRowData = startRowHeader + 1;
                        }
                        data.Values = new List<string>();
                        for (int z = startRowData; z < rows; z++)
                        {
                            string value = worksheet.Cells[z, i].Value?.ToString() ?? string.Empty;
                            data.Values.Add(value);
                        }
                    }
                    else
                    {

                        addressCellMerged = worksheet.MergedCells[startRowHeader, i];
                        /// Get Parameter
                        data.Parameter = worksheet.Cells[rowParameter, i].Value.ToString();

                        /// Get Component
                        if ((worksheet.Cells[startRowHeader, i].Value != null && worksheet.Cells[startRowHeader, i].Merge == true) || worksheet.Cells[startRowHeader, i].Merge == false)
                        {
                            data.Component = worksheet.Cells[startRowHeader, i].Value.ToString();
                            componentMerged = data.Component;
                            worksheet.GetMergeCellId(startRowHeader, i);
                        }
                        else
                        {
                            if (addressCellMerged.Contains(data.Component = worksheet.Cells[startRowHeader, i].Address))
                            {
                                data.Component = componentMerged;
                            }
                        }
                        /// Get Values
                        data.Values = new List<string>();
                        for (int j = startRowData; j <= rows; j++)
                        {
                            data.Values.Add(worksheet.Cells[j, i].Value != null ? worksheet.Cells[j, i].Value.ToString() : string.Empty);
                        }
                    }

                    dataCollection.Add(data);

                }
                return dataCollection;

            }
            return null;

        }
        public bool ExportToXML(DataCollection dataCollection, string filePath)
        {
            var sts = new XmlWriterSettings()
            {
                Indent = true,
            };

            using (var writer = XmlWriter.Create(filePath, sts))
            {
                writer.WriteStartDocument();

                var dataDistincCoponent = dataCollection.DataList.Select(x => x.Component).Distinct().ToList(); // Lấy các header có trong ws
                int countDataInComponentEmpty = dataCollection.DataList.Count(x => string.IsNullOrEmpty(x.Component)); //Đêm số lượng header là empty trong ws
                int numberCol = dataCollection.DataList.Count(x => string.IsNullOrEmpty(x.Parameter));  // Đếm số lượng cột của giá trị tương ứng vs các header
                int maxRowOfValue = dataCollection.DataList.Max(x => x.Values.Count); // data max row
                var dataHasMaxRowValue = dataCollection.DataList.FirstOrDefault(x => x.Values.Count == maxRowOfValue);//  

                writer.WriteStartElement("Root");
                /// ws no header
                if (1 == dataDistincCoponent.Count && countDataInComponentEmpty == numberCol)
                {
                    for (int i = 0; i < dataHasMaxRowValue.Values.Count; i++)
                    {
                        writer.WriteStartElement("Record");
                        writer.WriteAttributeString("id", (i + 1).ToString());
                        for (int j = 0; j < dataCollection.DataList.Count; j++)
                        {
                            writer.WriteStartElement("col");
                            writer.WriteAttributeString("id", (j + 1).ToString());
                            writer.WriteString(dataCollection.DataList[j].Values[i]);
                            writer.WriteEndElement();

                        }
                        writer.WriteEndElement();
                    }
                } // have header
                else 
                {
                    for (int i = 0; i < dataHasMaxRowValue.Values.Count; i++)
                    {
                        writer.WriteStartElement("Record");
                        writer.WriteAttributeString("id", (i + 1).ToString());
                        for (int j = 0; j < dataDistincCoponent.Count; j++) 
                        {
                            var dataWithComponent = dataCollection.DataList.Where(x => x.Component.Equals(dataDistincCoponent[j])).ToList();
                            if (dataDistincCoponent.Count == numberCol) // Only Header
                            {
                                for (int z = 0; z < dataWithComponent.Count; z++)
                                {
                                    if (string.IsNullOrEmpty(dataWithComponent[z].Component))
                                    {
                                        writer.WriteStartElement(RemoveSignalUnicodeCharacters("Header"+j));
                                        writer.WriteString(dataWithComponent[z].Values[i]);
                                        writer.WriteEndElement();
                                    }
                                    else
                                    {
                                        writer.WriteStartElement(RemoveSignalUnicodeCharacters(dataWithComponent[z].Component));
                                        writer.WriteString(dataWithComponent[z].Values[i]);
                                        writer.WriteEndElement();
                                    }
                                }
                            }
                            else
                            {
                                writer.WriteStartElement(RemoveSignalUnicodeCharacters(dataDistincCoponent[j]));
                                for (int z = 0; z < dataWithComponent.Count; z++) //Header merged row 2 
                                {
                                    writer.WriteStartElement(RemoveSignalUnicodeCharacters(dataWithComponent[z].Parameter));
                                    writer.WriteString(dataWithComponent[z].Values[i]);
                                    writer.WriteEndElement();
                                }
                                writer.WriteEndElement();
                            }
                        }
                        writer.WriteEndElement();
                    }
                }
                writer.WriteEndElement();

                writer.WriteEndDocument();
                writer.Flush();
                writer.Close();
            }
            return true;
        }


        public string RemoveSignalUnicodeCharacters(string str)
        {
            string[] VietNamChar = new string[] {"aAeEoOuUiIdDyY", "áàạảãâấầậẩẫăắằặẳẵ", "ÁÀẠẢÃÂẤẦẬẨẪĂẮẰẶẲẴ", "éèẹẻẽêếềệểễ", "ÉÈẸẺẼÊẾỀỆỂỄ", "óòọỏõôốồộổỗơớờợởỡ", "ÓÒỌỎÕÔỐỒỘỔỖƠỚỜỢỞỠ", "úùụủũưứừựửữ", "ÚÙỤỦŨƯỨỪỰỬỮ", "íìịỉĩ", "ÍÌỊỈĨ", "đ", "Đ", "ýỳỵỷỹ", "ÝỲỴỶỸ" };
            if (string.IsNullOrWhiteSpace(str))
            {
                return string.Empty;
            }
            string result = str;
            result = string.Concat(result.Normalize(NormalizationForm.FormD).Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark));
            //Thay thế và lọc dấu từng char      
            for (int i = 1; i < VietNamChar.Length; i++)
            {
                for (int j = 0; j < VietNamChar[i].Length; j++)
                    result = result.Replace(VietNamChar[i][j], VietNamChar[0][i - 1]);
            }
            result = result.Replace(" ","") ;
            result = Regex.Replace(result, @"[^\u0000-\u007F]+", string.Empty);
            result = Regex.Replace(result, @"[!@#$%^&*()_+\-=\[\]{};\':\\\|,.<>\/?]", string.Empty);
            return result;
        }
    }
}
