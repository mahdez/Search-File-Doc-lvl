using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SearchFileDocument
{
    public partial class Form1 : Form
    {
        public static int i = 1;
        public bool CheckShow = false;
        public string ExportPath = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            textBox1.Text = "Searching...";
            ChooseFolderBrowse();

        }

        public void ChooseFolderBrowse()
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
            }

        }
        public void ChooseFolderExport()
        {
            if (folderBrowserDialogExport.ShowDialog() == DialogResult.OK)
            {
                ExportPath = folderBrowserDialogExport.SelectedPath;
            }
        }

        private void btnShow_Click(object sender, EventArgs e)
        {
            
            ClearData();
            var dir = textBox1.Text;
            PrintDirectory(dir, 99, null);
            lblTotal.Text = "Total " + (i - 1);
        }

        private void ClearData()
        {
            if (i > 1)
            {
                i = 1;
            }
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
        }

        public void PrintDirectory(string directory, int lvl, string[] excludedFolders )
        {
            excludedFolders = excludedFolders ?? new string[0];
            if (directory != null && directory != "" && directory != "Searching...")
            {
                foreach (string f in Directory.GetFiles(directory))
                {
                    this.dataGridView1.Rows.Add(i++, Path.GetFileName(f), directory);
                }
                foreach (string d in Directory.GetDirectories(directory))
                {
                    this.dataGridView1.Rows.Add(i++, Path.GetFileName(d), directory);

                    if (lvl > 0 && Array.IndexOf(excludedFolders, Path.GetFileName(d)) < 0)
                    {
                        PrintDirectory(d, lvl - 1, excludedFolders);
                    }
                    
                }
                CheckShow = true;
            }
            else
            {
                string message = "Please Check Path File";
                MessageBox.Show(message);
            }
            
        }

        public void ExportExcel(string directory)
        {
            if (CheckShow == false)
            {
                string message = "Please Click Show First";
                MessageBox.Show(message);
            }
            else if (directory != null && directory != "" && directory != "Searching...")
            {
                ExcelPackage excel = new ExcelPackage();
                var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

                workSheet.TabColor = System.Drawing.Color.Black;
                workSheet.DefaultRowHeight = 12;

                workSheet.Row(1).Height = 20;
                workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Row(1).Style.Font.Bold = true;

                //workSheet.Cells[1, 1].Value = "No.";
                workSheet.Cells[1, 1].Value = "FileName";
                workSheet.Cells[1, 2].Value = "Path";

                int recordIndex = 2;
                //var data = (dataGridView1.DataSource);
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        string value = cell.Value.ToString();
                        if (cell.ColumnIndex == 1)
                        {
                            workSheet.Cells[recordIndex, 1].Value = value;

                        }
                        else if (cell.ColumnIndex == 2)
                        {
                            workSheet.Cells[recordIndex, 2].Value = value;
                            recordIndex++;
                        }

                    }
                }

                workSheet.Column(1).AutoFit();
                workSheet.Column(2).AutoFit();
                //workSheet.Column(3).AutoFit(); folderBrowserDialogExport
                //    FileStream objFileStrm = File.Create(p_strPath);
                //    objFileStrm.Close();
                string p_strPath = ExportPath + "\\NamePath.xlsx";
                string extension = Path.GetExtension(p_strPath);

                int l = 0;
                while (File.Exists(p_strPath))
                {
                    if (l == 0)
                        p_strPath = p_strPath.Replace(extension, "(" + ++l + ")" + extension);
                    else
                        p_strPath = p_strPath.Replace("(" + l + ")" + extension, "(" + ++l + ")" + extension);
                }

                File.WriteAllBytes(p_strPath, excel.GetAsByteArray());

                excel.Dispose();
                Console.Read();
                
                string message = "Download Complete";
                MessageBox.Show(message);
            }
            else
            {
                string message = "Please Check Path File";
                MessageBox.Show(message);
            }

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            ChooseFolderExport();
            ExportExcel(textBox1.Text);

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}

