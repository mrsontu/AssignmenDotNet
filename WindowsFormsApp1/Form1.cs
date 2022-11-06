using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using WindowsFormsApp1.Entities;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowseFile_Click(object sender, EventArgs e)
        {
            opn_FileDialog.ShowDialog();
        }

        private void btn_XMLFile_Click(object sender, EventArgs e)
        {
            ofdXMLFile.ShowDialog();
        }

        private void ofdXMLFile_FileOk(object sender, CancelEventArgs e)
        {
            txt_XMLFilePath.Text = ofdXMLFile.FileName;
            dgv_FileData.Enabled = true;
            this.BindingDataToGridView(this, new EventArgs());
        }

        private void opn_FileDialog_FileOk(object sender, CancelEventArgs e)
        {
            txtFilePath.Text = opn_FileDialog.FileName;
        }

        private void txtDiagLogOK(object sender, CancelEventArgs e)
        {
            textBox1.Text = txtFile.FileName;
        }

        private async void btn_ExcelFromData_Click(object sender, EventArgs e)
        {
            try
            {

                //Dùng file path làm key để lock file khi có những thằng cùng đọc 1 file
                string fileName = $"ExportXmlFromExcel{DateTime.Now.ToString("yyyyMMddhhmmss")}.xml";
                var path = Environment.CurrentDirectory + "//Files//ExportXML//" + fileName;
                if (string.IsNullOrEmpty(this.ExcelFilePath))
                {
                    MessageBox.Show("Please choose excel file import", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    int startHeader = 0;
                    if (string.IsNullOrEmpty(txtStartHeaderXML.Text.Trim()) || int.TryParse(txtStartHeaderXML.Text.Trim(), out startHeader) == false)
                    {
                        MessageBox.Show("Please enter  start row header", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    this.Enabled = false;
                    await Task.Run(() =>
                    {
                        lock (this.ExcelFilePath)
                        {
                            DataCollection dataCollection = new DataCollection();
                            //dataCollection = dataCollection.ReadDataFromExcel(this.ExcelFilePath, startHeader);
                            int endOfHeader;
                            var headers = dataCollection.GetHeadersFromExcel(this.ExcelFilePath, out endOfHeader);
                            var cells = dataCollection.GetCellsFromExcel(this.ExcelFilePath, endOfHeader);
                            dataCollection.ExportToXML(path, cells, headers);
                            //dataCollection.ExportToXML(dataCollection, path);
                        }
                    });
                    this.Enabled = true;
                    MessageBox.Show($"Export data to xml successfully  {path}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

            }
            catch (Exception ex)
            {
                this.Enabled = true;
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private async void btn_GenerateData_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(this.TextPath))
                {
                    MessageBox.Show("Please choose excel file import", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    int startRowHeader = 0;
                    int startRowData = 0;
                    if (string.IsNullOrEmpty(txtStartRow.Text.Trim()) || int.TryParse(txtStartRow.Text.Trim(), out startRowData) == false)
                    {
                        MessageBox.Show("Please enter  start row data", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    if (string.IsNullOrEmpty(txtStartHeader.Text.Trim()) || int.TryParse(txtStartHeader.Text.Trim(), out startRowHeader) == false)
                    {
                        MessageBox.Show("Please enter  start row header", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    await Task.Run(() =>
                    {
                        lock (this.TextPath)
                        {
                            string fileName = $"ExportTextToExcel{DateTime.Now.ToString("yyyyMMddhhmmss")}.xlsx";
                            string pathExport = System.IO.Directory.GetCurrentDirectory() + "//Files//ExportExcel//" + fileName;
                            DataFromText dataFromText = new DataFromText();
                            dataFromText.ReadDataFromText(startRowData, startRowHeader, this.TextPath, ',');
                            if (dataFromText != null)
                            {
                                dataFromText.ExportToExcel(pathExport);
                            }
                            MessageBox.Show($"Export data to excel successfully {this.TextPath}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private async void BindingDataToGridView(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(this.XmlFilePath))
                {
                    MessageBox.Show("Please choose import xml", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                  
                            System.IO.StringWriter swXML = new System.IO.StringWriter();
                            dataSetXml.ReadXml(this.XmlFilePath);
                            this.DataGridViewControl.DataSource = dataSetXml;
                            this.DataGridViewControl.DataMember = dataSetXml.Tables[0].TableName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void SaveXML_Click(object sender, EventArgs e)
        {
            try
            {
                string path = this.XmlFilePath;
                FileInfo fileInfo1 = new FileInfo(path);
                int indexExtenstion = fileInfo1.Name.LastIndexOf('.');
                string fileExportName = $"{fileInfo1.Name.Substring(0, indexExtenstion + 1)}-new{fileInfo1.Extension}";
                string filePath = Path.Combine(fileInfo1.Directory.FullName, fileExportName);
                DataSet dataSetSave = (DataSet)this.DataGridViewControl.DataSource;
                dataSetSave.WriteXml(filePath);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.txtFile.ShowDialog();
        }

        private async void button2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                DataCollection dataCollection = new DataCollection();
                string fileName = $"DataGeneateToExcel{DateTime.Now.ToString("yyyyMMddhhmmss")}.xlsx";
                string pathExport = System.IO.Directory.GetCurrentDirectory() + "//Files//" + fileName;
                dataCollection = dataCollection.GetDataFromTxt();
                dataCollection.ExportToExcel(dataCollection, pathExport);
                MessageBox.Show($"Export data to excel successfully {pathExport}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
