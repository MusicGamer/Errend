using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlServerCe;

namespace Errend
{
    public partial class Form3 : Form
    {
        static string sqlConnection = @"Data Source = " + Properties.Settings.Default["DBPath"].ToString() + "; Persist Security Info=False";
        SqlCeDataAdapter da;
        DataSet ds;
        Form f = new Form1();
        SqlCeConnection conn = new SqlCeConnection(sqlConnection);
        bool dirtyData = false;
        public Form3()
        {
            InitializeComponent();
        }

        private void ShowData(string nameTable)
        {
            try
            {
                ds = new DataSet();
                string query;
                conn.Open();
                switch (nameTable)
                {
                    case "Cargo":
                        query = "select * from Cargo";
                        da = new SqlCeDataAdapter(query, conn);
                        da.Fill(ds);
                        dataGridView1.DataSource = ds.Tables[0];
                        break;
                    case "CountryPort":
                        query = "select * from CountryPort";
                        da = new SqlCeDataAdapter(query, conn);
                        da.Fill(ds);
                        dataGridView1.DataSource = ds.Tables[0];
                        break;
                    case "Lines":
                        query = "select * from Lines";
                        da = new SqlCeDataAdapter(query, conn);
                        da.Fill(ds);
                        dataGridView1.DataSource = ds.Tables[0];
                        break;
                    case "PortT":
                        query = "select * from PortT";
                        da = new SqlCeDataAdapter(query, conn);
                        da.Fill(ds);
                        dataGridView1.DataSource = ds.Tables[0];
                        break;
                    case "Receiver":
                        query = "select * from Receiver";
                        da = new SqlCeDataAdapter(query, conn);
                        da.Fill(ds);
                        dataGridView1.DataSource = ds.Tables[0];
                        break;
                    case "Sender":
                        query = "select * from Sender";
                        da = new SqlCeDataAdapter(query, conn);
                        da.Fill(ds);
                        dataGridView1.DataSource = ds.Tables[0];
                        break;
                    case "Vessels":
                        query = "select * from Vessels";
                        da = new SqlCeDataAdapter(query, conn);
                        da.Fill(ds);
                        dataGridView1.DataSource = ds.Tables[0];
                        break;
                    default:
                        break;
                }
                conn.Close();
            }
            catch (Exception)
            {

            }
        }

        private void UpdateDataBase()
        {
            conn.Open();
            SqlCeCommandBuilder cb = new SqlCeCommandBuilder(da);
            da.Update(ds.Tables[0]);
            conn.Close();
            dirtyData = false;
        }

        private void update_Click(object sender, EventArgs e)
        {
            UpdateDataBase();
        }

        private void viewLines_Click(object sender, EventArgs e)
        {
            AskAboutUpdate();
            ShowData("Lines");
        }

        private void viewSender_Click(object sender, EventArgs e)
        {
            AskAboutUpdate();
            ShowData("Sender");
        }

        private void viewReceiver_Click(object sender, EventArgs e)
        {
            AskAboutUpdate();
            ShowData("Receiver");
        }

        private void viewVessels_Click(object sender, EventArgs e)
        {
            AskAboutUpdate();
            ShowData("Vessels");
        }

        private void viewPortT_Click(object sender, EventArgs e)
        {
            AskAboutUpdate();
            ShowData("PortT");
        }

        private void viewCountryPort_Click(object sender, EventArgs e)
        {
            AskAboutUpdate();
            ShowData("CountryPort");
        }

        private void viewCargo_Click(object sender, EventArgs e)
        {
            AskAboutUpdate();
            ShowData("Cargo");
        }

        private void AskAboutUpdate()
        {
            if (dirtyData)
            {
                if (MessageBox.Show("Сохранить изменения?", "DataBase", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    UpdateDataBase();
                }
                else
                {
                    dirtyData = false;
                }
            }
        }

        private void dataGridView1_CelEndEdite(object sender, DataGridViewCellEventArgs e)
        {
            dirtyData = true;
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            AskAboutUpdate();
            e.Cancel = false;
        }
    }
}
