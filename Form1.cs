using System;
using System.Data;
using System.IO;
using System.Linq;
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

        private void label1_Click(object sender, EventArgs e) { }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnGenerateXML.Enabled = false;
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            try
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

                        btnGenerateXML.Enabled = true;
                        MessageBox.Show("Excel file loaded successfully.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("Error uploading Excel file.", ex);
                MessageBox.Show("An error occurred while uploading the file. Check error log for details.");
            }
        }


        private void btnGenerateXML_Click(object sender, EventArgs e)
        {
            try
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
                        btnGenerateXML.Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("Error generating XML file.", ex);
                MessageBox.Show("An error occurred while generating the XML file. Check error log for details.");
            }
        }


        private void ConvertToXml(string filePath)
        {
            DataTable dataTable = excelData.Tables[0];

            using (XmlWriter writer = XmlWriter.Create(filePath, new XmlWriterSettings { Indent = true, IndentChars = "  " }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("transactions");

                foreach (DataRow row in dataTable.Rows)
                {
                    writer.WriteStartElement("transaction");

                    WriteElementIfNotEmpty(writer, "transactionnumber", ValidateString(row["Trans Reference"]));
                    WriteElementIfNotEmpty(writer, "transaction_description", FormatDescription(ValidateString(row["TRANS DESCRIPTION"])));
                    WriteElementIfNotEmpty(writer, "date_transaction", FormatDate(row["TRANSACTION DATE"]));
                    WriteElementIfNotEmpty(writer, "amount_local", ValidateString(row["AMOUNT LCY"]));

                    writer.WriteStartElement("t_from_my_client");
                    WriteElementIfNotEmpty(writer, "from_funds_code", ValidateEnumeration(row["TRANS DESCRIPTION"], new[] { "AAT", "BC01", "BST", "C", "CD1", "CDFC", "CDOB", "COB", "CTBB", "CW1", "CWFC", "E", "H", "IFM", "INT", "INTC", "LP", "MM", "MR", "OCCD", "OESP", "OFM", "PCC", "STO", "TFA", "TFDB", "TTA", "TTDB" }));

                    writer.WriteStartElement("from_person");
                    WriteElementIfNotEmpty(writer, "gender", MapGender(row["GENDER"]?.ToString()));
                    (string firstName, string lastName) = SplitName(ValidateString(row["TRANSACTING PERSON"]));
                    WriteElementIfNotEmpty(writer, "first_name", firstName);
                    WriteElementIfNotEmpty(writer, "last_name", lastName);
                    WriteElementIfNotEmpty(writer, "birthdate", FormatDate(row["CONDUCTOR DOB"]));
                    WriteElementIfNotEmpty(writer, "id_number", ValidateString(row["CONDUCTOR ID NUMBER"]));
                    WriteElementIfNotEmpty(writer, "nationality1", ValidateString(row["CONDUCTOR NATIONALITY"]));
                    WriteElementIfNotEmpty(writer, "occupation", ValidateString(row["CONDUCTOR OCCUPATION"]));
                    writer.WriteEndElement();

                    WriteElementIfNotEmpty(writer, "from_country", ValidateString(row["SIGNATORY NATIONALITY"]));
                    writer.WriteEndElement();

                    writer.WriteStartElement("t_to_my_client");
                    WriteElementIfNotEmpty(writer, "to_funds_code", ValidateEnumeration(row["TRANS DESCRIPTION"], new[] { "AAT", "BC01", "BST", "C", "CD1", "CDFC", "CDOB", "COB", "CTBB", "CW1", "CWFC", "E", "H", "IFM", "INT", "INTC", "LP", "MM", "MR", "OCCD", "OESP", "OFM", "PCC", "STO", "TFA", "TFDB", "TTA", "TTDB" }));

                    writer.WriteStartElement("to_account");
                    WriteElementIfNotEmpty(writer, "institution_name", ValidateMinLength(row["Institution Name"], 1));
                    WriteElementIfNotEmpty(writer, "swift", ValidateMinLength(row["SWIFT"], 1));
                    WriteElementIfNotEmpty(writer, "account", ValidateString(row["ACCOUNT NUMBER"]));
                    WriteElementIfNotEmpty(writer, "account_name", ValidateString(row["CUSTOMER NAME"]));
                    WriteElementIfNotEmpty(writer, "personal_account_type", ValidateString(row["PERSONAL ACCOUNT TYPE"]));

                    writer.WriteStartElement("signatory");
                    writer.WriteStartElement("t_person");
                    WriteElementIfNotEmpty(writer, "gender", MapGender(row["GENDER"]?.ToString()));
                    (firstName, lastName) = SplitName(ValidateString(row["CUSTOMER NAME"]));
                    WriteElementIfNotEmpty(writer, "first_name", firstName);
                    WriteElementIfNotEmpty(writer, "last_name", lastName);
                    WriteElementIfNotEmpty(writer, "birthdate", FormatDate(row["CONDUCTOR DOB"]));
                    WriteElementIfNotEmpty(writer, "legal_doc_name", ValidateString(row["LEGAL DOC NAME"]));
                    WriteElementIfNotEmpty(writer, "id_number", ValidateString(row["SIGNATORY PASSPORT ID"]));
                    WriteElementIfNotEmpty(writer, "occupation", ValidateString(row["SIGNATORY OCCUPATION"]));
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    WriteElementIfNotEmpty(writer, "opened", FormatDate(row["ISSUE DATE"]));
                    WriteElementIfNotEmpty(writer, "status_code", ValidateString(row["STATUS CODE"]));
                    writer.WriteEndElement();

                    WriteElementIfNotEmpty(writer, "to_country", ValidateString(row["SIGNATORY NATIONALITY"]));
                    writer.WriteEndElement();
                    WriteElementIfNotEmpty(writer, "comments", ValidateString(row["COMMENT"]));

                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
            
        }

        private void WriteElementIfNotEmpty(XmlWriter writer, string tagName, string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                writer.WriteElementString(tagName, value);
            }
        }

        private string ValidateString(object value)
        {
            return value?.ToString().Trim() ?? "";
        }

        private string FormatDate(object dateValue)
        {
            if (dateValue == null || string.IsNullOrWhiteSpace(dateValue.ToString()))
                return "";
            if (DateTime.TryParse(dateValue.ToString(), out DateTime date))
                return date.ToString("yyyy-MM-ddTHH:mm:ss");
            return "";
        }

        private string ValidateEnumeration(object value, string[] validValues)
        {
            string strValue = value?.ToString().Trim() ?? "";
            return validValues.Contains(strValue) ? strValue : "";
        }

        private string ValidateMinLength(object value, int minLength)
        {
            string strValue = value?.ToString().Trim() ?? "";
            return strValue.Length >= minLength ? strValue : "";
        }

        private string MapGender(string gender)
        {
            if (string.IsNullOrWhiteSpace(gender))
                return "";
            return gender.ToLower() == "male" ? "M" : gender.ToLower() == "female" ? "F" : "";
        }

        private string FormatDescription(string description)
        {
            if (string.IsNullOrWhiteSpace(description)) return "";
            description = description.ToLower();
            if (description.Contains("withdrawal"))
                return "cash withdrawal";
            else if (description.Contains("deposit"))
                return "cash deposit";
            return description;
        }

        private (string firstName, string lastName) SplitName(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return ("", "");
            var names = fullName.Split(' ');
            return (names.FirstOrDefault() ?? "", names.Length > 1 ? string.Join(" ", names.Skip(1)) : "");
        }
        private void LogError(string message, Exception ex = null)
        {
            string logFilePath = Path.Combine(Application.StartupPath, "error_log.txt");
            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                writer.WriteLine("----- Error Log -----");
                writer.WriteLine("Date/Time: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                writer.WriteLine("Message: " + message);
                if (ex != null)
                {
                    writer.WriteLine("Exception: " + ex.Message);
                    writer.WriteLine("Stack Trace: " + ex.StackTrace);
                }
                writer.WriteLine("---------------------");
                writer.WriteLine();
            }
        }
    }
}
