using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Errend
{
    public partial class Form2 : Form
    {
        bool dirtyData;
        public Form2()
        {
            InitializeComponent();
            textBox1.Text = Properties.Settings.Default["ErrentTable"].ToString();
            textBox2.Text = Properties.Settings.Default["ContainerTable"].ToString();
            textBox3.Text = Properties.Settings.Default["SavingPath"].ToString();
            textBox4.Text = Properties.Settings.Default["DBPath"].ToString();
            dirtyData = false;
        }

        private void save_Click(object sender, EventArgs e)
        {
            SaveSettings();
        }

        private void Path_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = fbd.SelectedPath;
            }
        }

        private void DBPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            if (fd.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = fd.FileName;
            }
        }

        private void SaveSettings()
        {
            Properties.Settings.Default["ErrentTable"] = textBox1.Text;
            Properties.Settings.Default["ContainerTable"] = textBox2.Text;
            Properties.Settings.Default["SavingPath"] = textBox3.Text;
            Properties.Settings.Default["DBPath"] = textBox4.Text;
            Properties.Settings.Default.Save();
            dirtyData = false;
        }

        private void AskAboutUpdate()
        {
            if (dirtyData)
            {
                if (MessageBox.Show("Сохранить изменения?", "Settings", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    SaveSettings();
                }
                else
                {
                    dirtyData = false;
                }
            }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            AskAboutUpdate();
            e.Cancel = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            dirtyData = true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            dirtyData = true;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            dirtyData = true;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            dirtyData = true;
        }
    }
}
