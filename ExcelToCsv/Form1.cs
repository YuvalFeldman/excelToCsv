using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace ExcelToCsv
{
    public partial class Form1 : Form
    {

        private OpenFileDialog _openFileDialog;
        private SaveFileDialog _saveFileDialog;
        private FolderBrowserDialog _folderBrowserDialog;

        public Form1()
        {
            InitializeComponent();
            InitializeReaders();
        }
        private void InitializeReaders()
        {
            _openFileDialog = new OpenFileDialog { Filter = @"Excel Files(.xlsx)|*.xlsx| Excel Files(.xls)|*.xls| Excel Files(.xlsm)|*.xlsm" };
            _saveFileDialog = new SaveFileDialog { Filter = @"CSV files (*.csv)|*.csv" };
            _folderBrowserDialog = new FolderBrowserDialog();
        }


        private void Convert_Click(object sender, EventArgs e)
        {
            string originUrl;
            string targetUrl;

            if (_openFileDialog.ShowDialog() == DialogResult.OK)
            {
                originUrl = _openFileDialog.FileName;
            }
            else
            {
                return;
            }

            if (_saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                targetUrl = _saveFileDialog.FileName;
            }
            else
            {
                return;
            }
            try
            {
                ConvertExcelToCsv(originUrl, targetUrl);
            }
            catch (Exception)
            {
                MessageBox.Show(@"There was an error in the request, could not complete the conversion");
                return;
            }

            MessageBox.Show(@"Action completed successfully");

        }

        private void BatchConvert_Click(object sender, EventArgs e)
        {
            string[] rawFiles;
            string saveDir;

            if (_folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                rawFiles = Directory.GetFiles(_folderBrowserDialog.SelectedPath);
            }
            else
            {
                return;
            }
            if (_folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                saveDir = _folderBrowserDialog.SelectedPath;
            }
            else
            {
                return;
            }
            try
            {
                foreach (var originUrl in rawFiles)
                {
                    var extension = Path.GetExtension(originUrl);
                    if (extension == null || (!extension.Equals(".xlsx") && !extension.Equals(".xls") && !extension.Equals(".xlsm"))) continue;
                    var targetUrl = saveDir + @"\" + Path.GetFileNameWithoutExtension(originUrl) + ".csv";
                    ConvertExcelToCsv(originUrl, targetUrl);
                }
            }
            catch (Exception)
            {
                MessageBox.Show(@"There was an error in the request, could not complete the conversion");
                return;
            }

            MessageBox.Show(@"Action completed successfully");
        }

        private void ConvertExcelToCsv(string excelFilePath, string csvOutputFile)
        {
            ValidateFileNonExistance(csvOutputFile);
            using (var p = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var sheet = p.Workbook.Worksheets[1];
                using (var writer = File.CreateText(csvOutputFile))
                {
                    var rowCount = sheet.Dimension.End.Row;
                    var columnCount = sheet.Dimension.End.Column;
                    for (var r = 1; r <= rowCount; r++)
                    {
                        for (var c = 1; c <= columnCount; c++)
                        {
                            writer.Write(sheet.Cells[r, c].Value);
                            writer.Write(",");
                        }
                        writer.WriteLine();
                    }
                }
            }
        }

        private void ValidateFileNonExistance(string csvOutputFile)
        {
            if (File.Exists(csvOutputFile)) File.Delete(csvOutputFile);
        }
    }
}
