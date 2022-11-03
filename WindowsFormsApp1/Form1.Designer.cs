using System.Windows.Forms;

namespace WindowsFormsApp1
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private GroupBox groupBox3;

        private Button btn_CreateXMLFile;

        private Label lblExcelFile;

        private TextBox txtFilePath;

        private Button btnBrowseFile;

        private GroupBox groupBox2;

        private Label label1;

        private Button btn_ExcelFromData;

        private GroupBox grbDataView;

        private Label label2;

        private TextBox txt_XMLFilePath;

        private Button btn_XMLFile;

        private DataGridView dgv_FileData;

        private OpenFileDialog opn_FileDialog;

        private OpenFileDialog ofdXMLFile;
        private OpenFileDialog txtFile;

        private Button btnSaveXml;

        public DataGridView DataGridViewControl => dgv_FileData;

        public string TextPath => txtFile.FileName;
        public string ExcelFilePath => opn_FileDialog.FileName;

        public string ExcelFileName => opn_FileDialog.SafeFileName;

        public string XmlFilePath => ofdXMLFile.FileName;

        public string XmlFileName => ofdXMLFile.SafeFileName;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtStartHeaderXML = new System.Windows.Forms.TextBox();
            this.btn_CreateXMLFile = new System.Windows.Forms.Button();
            this.lblExcelFile = new System.Windows.Forms.Label();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnBrowseFile = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtStartRow = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtStartHeader = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_ExcelFromData = new System.Windows.Forms.Button();
            this.grbDataView = new System.Windows.Forms.GroupBox();
            this.btnSaveXml = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_XMLFilePath = new System.Windows.Forms.TextBox();
            this.btn_XMLFile = new System.Windows.Forms.Button();
            this.dgv_FileData = new System.Windows.Forms.DataGridView();
            this.opn_FileDialog = new System.Windows.Forms.OpenFileDialog();
            this.txtFile = new System.Windows.Forms.OpenFileDialog();
            this.ofdXMLFile = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.button5 = new System.Windows.Forms.Button();
            this.dataSetXml = new System.Data.DataSet();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.grbDataView.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_FileData)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataSetXml)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.txtStartHeaderXML);
            this.groupBox3.Controls.Add(this.btn_CreateXMLFile);
            this.groupBox3.Controls.Add(this.lblExcelFile);
            this.groupBox3.Controls.Add(this.txtFilePath);
            this.groupBox3.Controls.Add(this.btnBrowseFile);
            this.groupBox3.Location = new System.Drawing.Point(659, 15);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(340, 151);
            this.groupBox3.TabIndex = 18;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Export Data From Excel";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 52);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(92, 13);
            this.label6.TabIndex = 18;
            this.label6.Text = "Start Row Header";
            // 
            // txtStartHeaderXML
            // 
            this.txtStartHeaderXML.Location = new System.Drawing.Point(110, 52);
            this.txtStartHeaderXML.Name = "txtStartHeaderXML";
            this.txtStartHeaderXML.Size = new System.Drawing.Size(169, 20);
            this.txtStartHeaderXML.TabIndex = 17;
            // 
            // btn_CreateXMLFile
            // 
            this.btn_CreateXMLFile.Location = new System.Drawing.Point(205, 107);
            this.btn_CreateXMLFile.Name = "btn_CreateXMLFile";
            this.btn_CreateXMLFile.Size = new System.Drawing.Size(106, 23);
            this.btn_CreateXMLFile.TabIndex = 2;
            this.btn_CreateXMLFile.Text = "Generate XML File";
            this.btn_CreateXMLFile.UseVisualStyleBackColor = true;
            this.btn_CreateXMLFile.Click += new System.EventHandler(this.btn_ExcelFromData_Click);
            // 
            // lblExcelFile
            // 
            this.lblExcelFile.AutoSize = true;
            this.lblExcelFile.Location = new System.Drawing.Point(10, 26);
            this.lblExcelFile.Name = "lblExcelFile";
            this.lblExcelFile.Size = new System.Drawing.Size(88, 13);
            this.lblExcelFile.TabIndex = 12;
            this.lblExcelFile.Text = "Select Excel file :";
            // 
            // txtFilePath
            // 
            this.txtFilePath.Enabled = false;
            this.txtFilePath.Location = new System.Drawing.Point(110, 23);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(169, 20);
            this.txtFilePath.TabIndex = 10;
            // 
            // btnBrowseFile
            // 
            this.btnBrowseFile.Location = new System.Drawing.Point(285, 21);
            this.btnBrowseFile.Name = "btnBrowseFile";
            this.btnBrowseFile.Size = new System.Drawing.Size(26, 23);
            this.btnBrowseFile.TabIndex = 11;
            this.btnBrowseFile.Text = "...";
            this.btnBrowseFile.UseVisualStyleBackColor = true;
            this.btnBrowseFile.Click += new System.EventHandler(this.btnBrowseFile_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.txtStartRow);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.txtStartHeader);
            this.groupBox2.Controls.Add(this.textBox1);
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.btn_ExcelFromData);
            this.groupBox2.Location = new System.Drawing.Point(232, 15);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(421, 151);
            this.groupBox2.TabIndex = 17;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Import Data";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(17, 78);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(80, 13);
            this.label5.TabIndex = 18;
            this.label5.Text = "Start Row Data";
            // 
            // txtStartRow
            // 
            this.txtStartRow.Location = new System.Drawing.Point(134, 78);
            this.txtStartRow.Name = "txtStartRow";
            this.txtStartRow.Size = new System.Drawing.Size(169, 20);
            this.txtStartRow.TabIndex = 17;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 47);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(92, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "Start Row Header";
            // 
            // txtStartHeader
            // 
            this.txtStartHeader.Location = new System.Drawing.Point(134, 47);
            this.txtStartHeader.Name = "txtStartHeader";
            this.txtStartHeader.Size = new System.Drawing.Size(169, 20);
            this.txtStartHeader.TabIndex = 15;
            // 
            // textBox1
            // 
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(134, 21);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(169, 20);
            this.textBox1.TabIndex = 13;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(309, 19);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(26, 23);
            this.button1.TabIndex = 14;
            this.button1.Text = "...";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 13);
            this.label1.TabIndex = 13;
            this.label1.Text = "Txt file path";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // btn_ExcelFromData
            // 
            this.btn_ExcelFromData.Location = new System.Drawing.Point(266, 122);
            this.btn_ExcelFromData.Name = "btn_ExcelFromData";
            this.btn_ExcelFromData.Size = new System.Drawing.Size(107, 23);
            this.btn_ExcelFromData.TabIndex = 0;
            this.btn_ExcelFromData.Text = "Generate Excel File";
            this.btn_ExcelFromData.UseVisualStyleBackColor = true;
            this.btn_ExcelFromData.Click += new System.EventHandler(this.btn_GenerateData_Click);
            // 
            // grbDataView
            // 
            this.grbDataView.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.grbDataView.Controls.Add(this.btnSaveXml);
            this.grbDataView.Controls.Add(this.label2);
            this.grbDataView.Controls.Add(this.txt_XMLFilePath);
            this.grbDataView.Controls.Add(this.btn_XMLFile);
            this.grbDataView.Controls.Add(this.dgv_FileData);
            this.grbDataView.Location = new System.Drawing.Point(12, 206);
            this.grbDataView.Name = "grbDataView";
            this.grbDataView.Size = new System.Drawing.Size(971, 362);
            this.grbDataView.TabIndex = 16;
            this.grbDataView.TabStop = false;
            this.grbDataView.Text = "Select Xml File and Bind DataGrid View";
            // 
            // btnSaveXml
            // 
            this.btnSaveXml.Location = new System.Drawing.Point(613, 268);
            this.btnSaveXml.Name = "btnSaveXml";
            this.btnSaveXml.Size = new System.Drawing.Size(78, 23);
            this.btnSaveXml.TabIndex = 16;
            this.btnSaveXml.Text = "Save XML";
            this.btnSaveXml.UseVisualStyleBackColor = true;
            this.btnSaveXml.Click += new System.EventHandler(this.SaveXML_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(87, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "Select XML File :";
            // 
            // txt_XMLFilePath
            // 
            this.txt_XMLFilePath.Enabled = false;
            this.txt_XMLFilePath.Location = new System.Drawing.Point(101, 37);
            this.txt_XMLFilePath.Name = "txt_XMLFilePath";
            this.txt_XMLFilePath.Size = new System.Drawing.Size(236, 20);
            this.txt_XMLFilePath.TabIndex = 13;
            // 
            // btn_XMLFile
            // 
            this.btn_XMLFile.Location = new System.Drawing.Point(343, 34);
            this.btn_XMLFile.Name = "btn_XMLFile";
            this.btn_XMLFile.Size = new System.Drawing.Size(26, 23);
            this.btn_XMLFile.TabIndex = 14;
            this.btn_XMLFile.Text = "...";
            this.btn_XMLFile.UseVisualStyleBackColor = true;
            this.btn_XMLFile.Click += new System.EventHandler(this.btn_XMLFile_Click);
            // 
            // dgv_FileData
            // 
            this.dgv_FileData.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dgv_FileData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_FileData.Location = new System.Drawing.Point(6, 76);
            this.dgv_FileData.Name = "dgv_FileData";
            this.dgv_FileData.Size = new System.Drawing.Size(685, 178);
            this.dgv_FileData.TabIndex = 0;
            // 
            // opn_FileDialog
            // 
            this.opn_FileDialog.Filter = "(*.xls, *.xlsx)|*.xlsx; *.xlsx";
            this.opn_FileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.opn_FileDialog_FileOk);
            // 
            // txtFile
            // 
            this.txtFile.Filter = "(*.txt, *.csv)|*.txt; *.csv";
            this.txtFile.FileOk += new System.ComponentModel.CancelEventHandler(this.txtDiagLogOK);
            // 
            // ofdXMLFile
            // 
            this.ofdXMLFile.Filter = "(*.xml, *.config)|*.xml; *.config";
            this.ofdXMLFile.FileOk += new System.ComponentModel.CancelEventHandler(this.ofdXMLFile_FileOk);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.button5);
            this.groupBox1.Location = new System.Drawing.Point(23, 15);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(208, 151);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Import Data";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(25, 49);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(177, 23);
            this.button2.TabIndex = 16;
            this.button2.Text = "Generate Excel File With Model ";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(365, 24);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(26, 23);
            this.button4.TabIndex = 14;
            this.button4.Text = "...";
            this.button4.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(46, 29);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(138, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Generate Excel From Data :";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(284, 78);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(107, 23);
            this.button5.TabIndex = 0;
            this.button5.Text = "Generate Excel File";
            this.button5.UseVisualStyleBackColor = true;
            // 
            // dataSetXml
            // 
            this.dataSetXml.DataSetName = "NewDataSet";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(995, 590);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.grbDataView);
            this.Name = "Form1";
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.grbDataView.ResumeLayout(false);
            this.grbDataView.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_FileData)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataSetXml)).EndInit();
            this.ResumeLayout(false);

        }



        #endregion

        private System.Data.DataSet dataSetXml;
        private TextBox textBox1;
        private Button button1;
        private GroupBox groupBox1;
        private Button button4;
        private Label label3;
        private Button button5;
        private Button button2;
        private Label label5;
        private TextBox txtStartRow;
        private Label label4;
        private TextBox txtStartHeader;
        private Label label6;
        private TextBox txtStartHeaderXML;
    }
}

