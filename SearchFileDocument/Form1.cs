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

namespace SearchFileDocument
{
    public partial class Form1 : Form
    {
        public static int i = 1;
        public Form1()
        {
            InitializeComponent();
            
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            textBox1.Text = "Searching...";
            ChooseFolder();

        }

        public void ChooseFolder()
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            ClearData();
            var dir = textBox1.Text;
            PrintDirectoryTree(dir, 99, null);
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

        public void PrintDirectoryTree(string directory, int lvl, string[] excludedFolders, string lvlSeperator = "" )
        {
            excludedFolders = excludedFolders ?? new string[0];
            if (directory != null && directory != "")
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
                        PrintDirectoryTree(d, lvl - 1, excludedFolders, lvlSeperator + "  ");
                    }
                }
            }
            else
            {
                string message = "Please Check Path File";
                MessageBox.Show(message);
            }
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}

