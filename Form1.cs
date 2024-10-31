using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using ExcelDataReader; 

namespace exceml
{
    public partial class Form1 : Form
    {
        private DataSet excelData;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnGenerateXML.Enabled = false; // Disable Generate button until file is uploaded
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
                openFileDialog.Title = "Select an Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            excelData = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = _ => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                        }
                    }

                    btnGenerateXML.Enabled = true; // Enable Generate XML button after file upload
                    MessageBox.Show("Excel file loaded successfully.");
                }
            }
        }

        private void btnGenerateXML_Click(object sender, EventArgs e)
        {
            if (excelData == null || excelData.Tables.Count == 0)
            {
                MessageBox.Show("No data found in the Excel file.");
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "XML Files|*.xml";
                saveFileDialog.Title = "Save XML File";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ConvertToXml(saveFileDialog.FileName);
                    MessageBox.Show("XML file generated successfully.");
                }
            }
        }

        private void ConvertToXml(string filePath)
        {
            DataTable dataTable = excelData.Tables[0]; // Get the first sheet
            using (XmlWriter writer = XmlWriter.Create(filePath, new XmlWriterSettings { Indent = true }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Workbook");

                foreach (DataRow row in dataTable.Rows)
                {
                    writer.WriteStartElement("Row");

                    foreach (DataColumn column in dataTable.Columns)
                    {
                        writer.WriteStartElement(column.ColumnName);
                        writer.WriteString(row[column].ToString());
                        writer.WriteEndElement();
                    }

                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
    }
}
