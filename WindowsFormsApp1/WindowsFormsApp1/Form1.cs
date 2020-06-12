using Aspose.Words;
using Aspose.Words.Replacing;
using System.Linq;
using Newtonsoft.Json;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string folderPath="";
        private void btnConvert_Click(object sender, EventArgs e)
        {
            if (txtFileExcel.Text == "đường dẫn file excel sẽ hiện tại đây"
                || txtFileWord.Text == "đường dẫn file excel sẽ hiện tại đây"
                || txtSave.Text == "đường dẫn file lưu sẽ hiện tại đây")
            {
                MessageBox.Show("Kiểm tra lại đường dẫn");
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("Bạn có chắc chắn chưa?", "Thực hiện chuyển đổi", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    
                        FileStream stream = File.Open(txtFileExcel.Text, FileMode.Open, FileAccess.Read);
                        IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        var result = excelReader.AsDataSet();

                        result.Tables
                                                .OfType<DataTable>()
                                                .Select(c => c.TableName)
                                                .ToArray();
                        var table = result.Tables[0];
                        var listField = result.Tables[0].Rows[0];

                        for (int i = 1; i < result.Tables[0].Rows.Count; i++)
                        {
                            var x = new ExpandoObject() as IDictionary<string, Object>;
                            for (int j = 0; j < result.Tables[0].Columns.Count; j++)
                            {
                                x.Add(listField[j].ToString(), table.Rows[i][j]);
                            }

                            // Check if file already exists. If yes, delete it.     
                            var fileName = folderPath + @"\" + i + "_" + table.Rows[i][0] + ".docx";
                            
                            if (File.Exists(fileName) == false)
                            {
                                // Create a new file     
                                FileStream fs = File.Create(fileName);
                                fs.Close();
                            }
                            var templateEngine = new swxben.docxtemplateengine.DocXTemplateEngine();

                            templateEngine.Process(txtFileWord.Text, fileName, x);
                        }

                    
                }
                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }
            }
        }

        private void btnFileWord_Click(object sender, EventArgs e)
        {
            var filePath = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word files (*.docx)|*.docx";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                filePath = openFileDialog.FileName;
                txtSave.Text = filePath;
                
            }
            if (filePath != "")
                txtFileWord.Text = filePath;
        }

        private void btnFileExcel_Click(object sender, EventArgs e)
        {
            var filePath = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                filePath = openFileDialog.FileName;

                //Read the contents of the file into a stream
                //var fileStream = openFileDialog.OpenFile();

                //using (StreamReader reader = new StreamReader(fileStream))
                //{
                //    fileContent = reader.ReadToEnd();
                //}
            }
            if (filePath != "")
                txtFileExcel.Text = filePath;
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                folderPath = folderBrowserDialog1.SelectedPath;
                txtSave.Text = folderPath;
            }
        }
    }
}
