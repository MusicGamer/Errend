using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Xceed.Words.NET;
using System.Diagnostics;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Errend.Core;
using Errend.Core.CTO;
using System.Globalization;
using System.Data.SqlServerCe;

namespace Errend
{
    public partial class Form1 : Form
    {
        static string[] proforma = File.ReadAllLines(@"SomeTXT.txt", Encoding.GetEncoding(1251));

        ParserWorker<string[]> parser;
        string[][] vessels;
        DataSet ds;
        GoogleXLS gXLS = new GoogleXLS();
        double weight = 0;
        bool invalidField = false;
        ValueRange vr = null;
        string dbPath = Properties.Settings.Default["DBPath"].ToString();
        string sqlConnection = @"Data Source = " + Properties.Settings.Default["DBPath"].ToString() + "; Persist Security Info=False";
        string[] months = new string[] { "января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря" };

        public Form1()
        {
            InitializeComponent();
            string commonFolder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            string myAppFolder = Path.Combine(commonFolder, "Errend");
            AppDomain.CurrentDomain.SetData("DataDirectory", myAppFolder);
            parser = new ParserWorker<string[]>(new CtoParser());
            parser.OnCompleted += Parser_OnCompleted;
            parser.OnNewData += Parser_OnNewData;
            parser.Settings = new CtoSettings();
            parser.Start();
            FillCombatBox();
            ContData();
            arena.Checked = true;
        }

        private void Parser_OnNewData(object arg1, string[] arg2)
        {
            vessels = new string[arg2.Length][];
            for (int i = 0; i < arg2.Length; i++)
            {
                vessels[i] = arg2[i].Split('\n');
                for (int j = 0; j < vessels[i].Length; j++)
                {
                    vessels[i][j] = vessels[i][j].Trim();
                }
            }

            for (int i = 0; i < vessels.Length; i++)
            {
                if (!vesselsComboBox.Items.Contains(vessels[i][2]))
                {
                    vesselsComboBox.Items.Add(vessels[i][2]);
                }

            }
        }

        private void Parser_OnCompleted(object obj)
        {
            //MessageBox.Show("OK");
        }

        private void ContData()
        {
            string[] contSize = new string[] { "20", "40", "45" };
            string[] contType = new string[] { "DV", "HC", "HP", "RF", "PL" };
            contSizeComboBox.Items.AddRange(contSize);
            contTypeComboBox.Items.AddRange(contType);
        }

        private void CreateDocument_Click(object sender, EventArgs e)
        {
            CheckFields();
            if (invalidField)
            {
                invalidField = false;
                return;
            }
            ChangeFieldsColor();
            CreateDoc(gXLS.ReadDataFromGoogleXML(firstPoint.Text, secondPont.Text));
        }

        private void CheckFields()
        {
            invalidField = false;
            if (String.IsNullOrEmpty(errendNumber.Text))
            {
                errendNumber.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(linesComboBox.Text))
            {
                linesComboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(senderComboBox.Text))
            {
                senderComboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(receiverComboBox.Text))
            {
                receiverComboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }

            if (String.IsNullOrEmpty(vesselsComboBox.Text))
            {
                vesselsComboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(voyageComboBox.Text))
            {
                voyageComboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(seTextBox.Text))
            {
                seTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(flagTextBox.Text))
            {
                flagTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(countryСomboBox.Text))
            {
                countryСomboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(potComboBox.Text))
            {
                potComboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(podComboBox.Text))
            {
                podComboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(cargoNameСomboBox.Text))
            {
                cargoNameСomboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(cargoCodComboBox.Text))
            {
                cargoCodComboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(contSizeComboBox.Text))
            {
                contSizeComboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(contTypeComboBox.Text))
            {
                contTypeComboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(senderNameTextBox.Text))
            {
                senderNameTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(senderAddressTextBox.Text))
            {
                senderAddressTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(senderCodTextBox.Text))
            {
                senderCodTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(receiverAddressTextBox.Text))
            {
                receiverAddressTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(lineInfTextBox.Text))
            {
                lineInfTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(firstPoint.Text))
            {
                firstPoint.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(secondPont.Text))
            {
                secondPont.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
        }

        private void ChangeFieldsColor()
        {
            errendNumber.BackColor = System.Drawing.Color.White;
            linesComboBox.BackColor = System.Drawing.Color.White;
            senderComboBox.BackColor = System.Drawing.Color.White;
            receiverComboBox.BackColor = System.Drawing.Color.White;
            vesselsComboBox.BackColor = System.Drawing.Color.White;
            voyageComboBox.BackColor = System.Drawing.Color.White;
            seTextBox.BackColor = System.Drawing.Color.White;
            flagTextBox.BackColor = System.Drawing.Color.White;
            countryСomboBox.BackColor = System.Drawing.Color.White;
            potComboBox.BackColor = System.Drawing.Color.White;
            podComboBox.BackColor = System.Drawing.Color.White;
            cargoNameСomboBox.BackColor = System.Drawing.Color.White;
            cargoCodComboBox.BackColor = System.Drawing.Color.White;
            contSizeComboBox.BackColor = System.Drawing.Color.White;
            contTypeComboBox.BackColor = System.Drawing.Color.White;
            senderNameTextBox.BackColor = System.Drawing.Color.White;
            senderAddressTextBox.BackColor = System.Drawing.Color.White;
            senderCodTextBox.BackColor = System.Drawing.Color.White;
            receiverAddressTextBox.BackColor = System.Drawing.Color.White;
            lineInfTextBox.BackColor = System.Drawing.Color.White;
            firstPoint.BackColor = System.Drawing.Color.White;
            secondPont.BackColor = System.Drawing.Color.White;
        }

        private void ClearFields()
        {
            errendNumber.Text = "";
            linesComboBox.SelectedIndex = -1;
            linesComboBox.Text = "";
            senderComboBox.SelectedIndex = -1;
            senderComboBox.Text = "";        
            receiverComboBox.DataSource = new DataView(ds.Tables["Receiver"]).ToTable(true, "Name");
            receiverComboBox.SelectedIndex = -1;
            countryСomboBox.SelectedIndex = -1;
            countryСomboBox.Text = "";
            vesselsComboBox.SelectedIndex = -1;
            potComboBox.SelectedIndex = -1;
            potComboBox.Text = "";
            podComboBox.SelectedIndex = -1;
            podComboBox.Text = "";
            cargoNameСomboBox.SelectedIndex = -1;
            cargoCodComboBox.SelectedIndex = cargoNameСomboBox.SelectedIndex;
            cargoNameСomboBox.Text = "";
            bookingTextBox.Text = "";
            contSizeComboBox.SelectedIndex = -1;
            contTypeComboBox.SelectedIndex = -1;
            firstPoint.Text = "";
            secondPont.Text = "";
        }

        private void CreateDoc(ValueRange vr)
        {
            weight = 0;
            bool secondPage = false;
            int amountContPage1 = 0;
            int amounrContPage2 = 0;
            for (int i = 0; i < vr.Values.Count; i++)
            {
                weight += double.Parse(vr.Values[i][5].ToString(), CultureInfo.CreateSpecificCulture("uk-UA"));
            }
            string strWeight = weight.ToString("#,#.##", CultureInfo.CreateSpecificCulture("uk-UA"));
            if (vr.Values.Count > 22)
            {
                secondPage = true;
                amountContPage1 = 22;
                amounrContPage2 = vr.Values.Count - 22;
            }
            else
            {
                amountContPage1 = vr.Values.Count;
            }
            Directory.CreateDirectory(@"" + Properties.Settings.Default["SavingPath"].ToString() + "\\" + linesComboBox.Text + "\\" + vesselsComboBox.Text);
            DocX document = DocX.Create(@"" + Properties.Settings.Default["SavingPath"].ToString() + "\\" + linesComboBox.Text + "\\" + vesselsComboBox.Text + "\\" + senderComboBox.Text + " - Прч. " + errendNumber.Text + " - " + vr.Values.Count.ToString() + " конт. - " + countryСomboBox.Text + ".docx");
            Xceed.Words.NET.Image img;
            if (arena.Checked == true)
            {
                img = document.AddImage(@"logo1.png");
                
            }
            else
            {
                img = document.AddImage(@"UGLv1.jpg");
            }
            Picture p = img.CreatePicture();
            p.Height = (int)(138 / 3.2);
            p.Width = (int)(2207 / 3.2);
            document.AddHeaders();
            Header header = document.Headers.Odd;
            header.InsertParagraph(" ", false).InsertPicture(p).FontSize(1);
            document.MarginHeader = 0;
            document.MarginFooter = 0;
            document.MarginTop = 20;
            document.MarginLeft = 70;
            document.MarginRight = 25;

            document.AddFooters();
            Footer footer = document.Footers.Odd;
            Table tFooter = footer.InsertTable(2, 4);
            tFooter.Rows[1].Height = 50;
            tFooter.Rows[0].Cells[0].Paragraphs.First().IndentationBefore = -2f;
            tFooter.Rows[0].Cells[0].MarginLeft = 50;
            tFooter.Rows[0].Cells[3].MarginRight = 10;
            tFooter.Rows[0].Cells[0].Width = 160;
            tFooter.Rows[0].Cells[1].Width = 150;
            tFooter.Rows[0].Cells[2].Width = 150;
            tFooter.Rows[0].Cells[3].Width = 160;
            tFooter.Rows[0].Cells[0].Paragraphs.First().Append("Экспедитор").Alignment = Alignment.center;
            tFooter.Rows[0].Cells[1].Paragraphs.First().Append("Судовой агент").Alignment = Alignment.center;
            tFooter.Rows[0].Cells[2].Paragraphs.First().Append("Линейный агент").Alignment = Alignment.center;
            tFooter.Rows[0].Cells[3].Paragraphs.First().Append("Таможня").Alignment = Alignment.center;
            tFooter.Rows[1].Cells[0].Paragraphs.First().Append("Экспедитор\n").Alignment = Alignment.left;
            tFooter.Rows[1].Cells[0].Paragraphs[0].Append("Глоба В. Л.\n").Alignment = Alignment.left;
            tFooter.Rows[1].Cells[0].Paragraphs[0].Append("Тел.050-341-89-12").Alignment = Alignment.left;
            tFooter.Rows[1].Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            foreach (var item in tFooter.Paragraphs)
            {
                item.Font("Times New Roman").FontSize(10).Bold();
            }
            document.MarginFooter = 15;

            Paragraph p1 = document.InsertParagraph();
            p1.Append(proforma[0]);
            if (arena.Checked == true)
            {
                p1.Append(proforma[1]).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(Environment.NewLine).Append(proforma[2]);
            }
            else
            {
                p1.Append(proforma[29]).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(Environment.NewLine).Append(proforma[30]);
            }
            Paragraph p2 = document.InsertParagraph();
            p2.Append(proforma[3] + " ").Append("«" + linesComboBox.Text + "»").UnderlineStyle(UnderlineStyle.singleLine).SpacingAfter(8).Alignment = Alignment.right;

            Paragraph p3 = document.InsertParagraph();
            DateTime date = DateTime.Now;
            p3.Append(proforma[4] + " " + errendNumber.Text + " от «" + date.Day.ToString("d2") + "» " + months[date.Month - 1] + " " + date.Year + " г." + Environment.NewLine + proforma[5])
                .Append(proforma[6]).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[7]).SpacingAfter(8).Alignment = Alignment.center;

            Paragraph p4 = document.InsertParagraph();
            p4.Append(proforma[8] + " ").Append(senderNameTextBox.Text.ToUpper() + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(senderAddressTextBox.Text.ToUpper() + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append("Код: " + senderCodTextBox.Text).UnderlineStyle(UnderlineStyle.singleLine)
                .SpacingAfter(8);

            Paragraph p5 = document.InsertParagraph();
            p5.Append(proforma[9] + " ").Append(receiverComboBox.Text.ToUpper() + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(receiverAddressTextBox.Text.ToUpper()).UnderlineStyle(UnderlineStyle.singleLine);

            Table t1 = document.AddTable(5, 2);
            Xceed.Words.NET.Border b = new Xceed.Words.NET.Border(Xceed.Words.NET.BorderStyle.Tcbs_none, BorderSize.five, 0, System.Drawing.Color.White);
            t1.SetBorder(TableBorderType.InsideH, b);
            t1.SetBorder(TableBorderType.InsideV, b);
            t1.SetBorder(TableBorderType.Bottom, b);
            t1.SetBorder(TableBorderType.Top, b);
            t1.SetBorder(TableBorderType.Left, b);
            t1.SetBorder(TableBorderType.Right, b);

            t1.Rows[0].Cells[0].Width = 285;
            t1.Rows[0].Cells[1].MarginRight = 230;
            t1.Rows[1].Cells[0].Paragraphs.First().Append(proforma[10] + " ").Append("«" + vesselsComboBox.Text.ToUpper() + "»").UnderlineStyle(UnderlineStyle.singleLine).SpacingAfter(8);
            t1.Rows[1].Cells[1].Paragraphs.First().Append(proforma[11] + " ").Append(voyageComboBox.Text.ToUpper()).UnderlineStyle(UnderlineStyle.singleLine).Append("  " + proforma[12] + " ").Append(seTextBox.Text.ToUpper()).UnderlineStyle(UnderlineStyle.singleLine).Append("  " + proforma[13] + " ").Append(flagTextBox.Text.ToUpper()).UnderlineStyle(UnderlineStyle.singleLine).SpacingAfter(8);
            t1.Rows[2].Cells[0].Paragraphs.First().Append(proforma[14] + " ").Append(proforma[15]).UnderlineStyle(UnderlineStyle.singleLine).SpacingAfter(8);
            t1.Rows[2].Cells[1].Paragraphs.First().Append(proforma[16] + " ").Append(podComboBox.Text.ToUpper()).UnderlineStyle(UnderlineStyle.singleLine).SpacingAfter(8);
            t1.Rows[3].Cells[0].Paragraphs.First().Append(proforma[17] + " ").Append(potComboBox.Text.ToUpper()).UnderlineStyle(UnderlineStyle.singleLine).SpacingAfter(8);
            t1.Rows[4].Cells[0].Paragraphs.First().Append(proforma[18] + " ").Append(cargoNameСomboBox.Text.ToUpper()).UnderlineStyle(UnderlineStyle.singleLine).SpacingAfter(8);
            t1.Rows[4].Cells[1].Paragraphs.First().Append(proforma[19] + " ").Append(cargoCodComboBox.Text).UnderlineStyle(UnderlineStyle.singleLine);
            document.InsertTable(t1);

            Paragraph p6 = document.InsertParagraph();
            p6.Append(proforma[20] + " ").Append(bookingTextBox.Text.ToUpper()).UnderlineStyle(UnderlineStyle.singleLine);

            Table t2 = document.AddTable(amountContPage1 + 1, 8);
            t2.Rows[0].Cells[0].MarginLeft = 40;
            t2.Rows[0].Cells[0].Paragraphs.First().Append("Номер контейнера").Alignment = Alignment.center;
            t2.Rows[0].Cells[0].Paragraphs.First().IndentationBefore = -1f;
            t2.Rows[0].Cells[1].Paragraphs.First().Append("Тип").Alignment = Alignment.center;
            t2.Rows[0].Cells[2].Paragraphs.First().Append("Кол-во Мест").Alignment = Alignment.center;
            t2.Rows[0].Cells[3].Paragraphs.First().Append("Вес груза нетто").Alignment = Alignment.center;
            t2.Rows[0].Cells[4].Paragraphs.First().Append("Вес груза брутто").Alignment = Alignment.center;
            t2.Rows[0].Cells[5].Paragraphs.First().Append("VGM перепро-веренный вес к-ра брутто**").Alignment = Alignment.center;
            t2.Rows[0].Cells[6].Paragraphs.First().Append("Пломбы").Alignment = Alignment.center;
            t2.Rows[0].Cells[6].Width = 350;
            t2.Rows[0].Cells[7].MarginRight = 10;
            t2.Rows[0].Cells[7].Paragraphs.First().Append("ГТД").Alignment = Alignment.center;
            for (int i = 0; i < amountContPage1; i++)
            {
                t2.Rows[i + 1].Cells[0].Paragraphs.First().Append(vr.Values[i][0].ToString()).Alignment = Alignment.center;
                t2.Rows[i + 1].Cells[1].Paragraphs.First().Append(contSizeComboBox.Text + contTypeComboBox.Text).Alignment = Alignment.center;
                t2.Rows[i + 1].Cells[2].Paragraphs.First().Append(vr.Values[i][13].ToString().ToUpper() == "НАВАЛ" || vr.Values[i][13].ToString().ToUpper() == "" ? "НАВАЛ" : vr.Values[i][13].ToString()).Alignment = Alignment.center;
                t2.Rows[i + 1].Cells[3].Paragraphs.First().Append(vr.Values[i][4].ToString()).Alignment = Alignment.center;
                t2.Rows[i + 1].Cells[4].Paragraphs.First().Append(vr.Values[i][5].ToString()).Alignment = Alignment.center;
                t2.Rows[i + 1].Cells[5].Paragraphs.First().Append(vr.Values[i][6].ToString()).Alignment = Alignment.center;

                if (vr.Values[i].Count == 20)
                {
                    t2.Rows[i + 1].Cells[6].Paragraphs.First().Append(vr.Values[i][19].ToString()).Alignment = Alignment.center;
                }
                else
                {
                    t2.Rows[i + 1].Cells[6].Paragraphs.First().Append("").Alignment = Alignment.center;
                }
                t2.Rows[i + 1].Cells[7].Paragraphs.First().Append(vr.Values[i][2].ToString()).Alignment = Alignment.center;
                if (vr.Values[i][2].ToString() == "")
                {
                    t2.Rows[0].Cells[7].Width = 500;
                }
                else
                {
                    t2.Rows[0].Cells[7].Width = 0;
                }
            }
            foreach (var item in t2.Rows[0].Cells)
            {
                item.VerticalAlignment = VerticalAlignment.Center;
            }
            document.InsertTable(t2);

            Paragraph p7 = document.InsertParagraph();
            p7.Append(proforma[21] + Environment.NewLine + proforma[22]).SpacingAfter(8).Alignment = Alignment.center;

            Paragraph p8 = document.InsertParagraph();
            p8.Append(proforma[23] + " " + vr.Values.Count + "x" + contSizeComboBox.Text + "` контейнер(ов). ").Append("ВЕС " + '\u2013' + " " + strWeight + " кг." + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(proforma[24]).Append(proforma[25]).UnderlineStyle(UnderlineStyle.singleLine).SetLineSpacing(LineSpacingType.Line, 1.7f);

            Paragraph p9 = document.InsertParagraph();
            p9.Append(proforma[26] + " ").Append("ПРР: " + lineInfTextBox.Text.ToUpper() + ", ГРН" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append("                            ");
            if (arena.Checked == true)
            {
                p9.Append(proforma[27]).UnderlineStyle(UnderlineStyle.singleLine);
            }
            else
            {
                p9.Append(proforma[31]).UnderlineStyle(UnderlineStyle.singleLine);
            }

            foreach (var item in document.Paragraphs)
            {
                item.Font("Times New Roman").FontSize(10).Bold();
            }
            document.Paragraphs[2].FontSize(11);

            if (secondPage)
            {
                p9.InsertPageBreakAfterSelf();
                Paragraph p10 = document.InsertParagraph();
                p10.Append(proforma[28] + " " + errendNumber.Text + " от «" + date.Day.ToString("d2") + "» " + months[date.Month - 1] + " " + date.Year + " г." + Environment.NewLine + proforma[5])
                    .Append(proforma[6]).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[7]).SpacingAfter(8).Alignment = Alignment.center;

                Table t3 = document.AddTable(amounrContPage2 + 1, 8);
                t3.Rows[0].Cells[0].MarginLeft = 40;
                t3.Rows[0].Cells[0].Paragraphs.First().Append("Номер контейнера").FontSize(10).Alignment = Alignment.center;
                t3.Rows[0].Cells[0].Paragraphs.First().IndentationBefore = -1f;
                t3.Rows[0].Cells[1].Paragraphs.First().Append("Тип").FontSize(10).Alignment = Alignment.center;
                t3.Rows[0].Cells[2].Paragraphs.First().Append("Кол-во Мест").FontSize(10).Alignment = Alignment.center;
                t3.Rows[0].Cells[3].Paragraphs.First().Append("Вес груза нетто").FontSize(10).Alignment = Alignment.center;
                t3.Rows[0].Cells[4].Paragraphs.First().Append("Вес груза брутто").FontSize(10).Alignment = Alignment.center;
                t3.Rows[0].Cells[5].Paragraphs.First().Append("VGM перепро-веренный вес к-ра брутто**").FontSize(10).Alignment = Alignment.center;
                t3.Rows[0].Cells[6].Paragraphs.First().Append("Пломбы").FontSize(10).Alignment = Alignment.center;
                t3.Rows[0].Cells[6].Width = 350;
                t3.Rows[0].Cells[7].MarginRight = 10;
                t3.Rows[0].Cells[7].Paragraphs.First().Append("ГТД").FontSize(10).Alignment = Alignment.center;
                for (int i = 0; i < amounrContPage2; i++)
                {
                    t3.Rows[i + 1].Cells[0].Paragraphs.First().Append(vr.Values[i + 22][0].ToString()).FontSize(10).Alignment = Alignment.center;
                    t3.Rows[i + 1].Cells[1].Paragraphs.First().Append(contSizeComboBox.Text + contTypeComboBox.Text).FontSize(10).Alignment = Alignment.center;
                    t3.Rows[i + 1].Cells[2].Paragraphs.First().Append(vr.Values[i + 22][13].ToString().ToUpper() == "НАВАЛ" || vr.Values[i + 22][13].ToString().ToUpper() == "" ? "НАВАЛ" : vr.Values[i + 22][13].ToString()).FontSize(10).Alignment = Alignment.center;
                    t3.Rows[i + 1].Cells[3].Paragraphs.First().Append(vr.Values[i + 22][4].ToString()).FontSize(10).Alignment = Alignment.center;
                    t3.Rows[i + 1].Cells[4].Paragraphs.First().Append(vr.Values[i + 22][5].ToString()).FontSize(10).Alignment = Alignment.center;
                    t3.Rows[i + 1].Cells[5].Paragraphs.First().Append(vr.Values[i + 22][6].ToString()).FontSize(10).Alignment = Alignment.center;
                    if (vr.Values[i].Count == 20)
                    {
                        t3.Rows[i + 1].Cells[6].Paragraphs.First().Append(vr.Values[i + 22][19].ToString()).FontSize(10).Alignment = Alignment.center;
                    }
                    else
                    {
                        t3.Rows[i + 1].Cells[6].Paragraphs.First().Append("").FontSize(10).Alignment = Alignment.center;
                    }
                    t3.Rows[i + 1].Cells[7].Paragraphs.First().Append(vr.Values[i + 22][2].ToString()).FontSize(10).Alignment = Alignment.center;
                    if (vr.Values[i + 22 ][2].ToString() == "")
                    {
                        t3.Rows[0].Cells[7].Width = 500;
                    }
                    else
                    {
                        t3.Rows[0].Cells[7].Width = 0;
                    }
                }
                foreach (var item in t3.Rows[0].Cells)
                {
                    item.VerticalAlignment = VerticalAlignment.Center;
                }
                document.InsertTable(t3);

                Paragraph p11 = document.InsertParagraph();
                p11.Append(proforma[21] + Environment.NewLine + proforma[22]).SpacingAfter(8).FontSize(10).Alignment = Alignment.center;

                Paragraph p12 = document.InsertParagraph();
                p12.Append(proforma[23] + " " + vr.Values.Count + "x" + contSizeComboBox.Text + "` контейнер(ов)." + " ").FontSize(10).Append("ВЕС " + '\u2013' + " " + strWeight + " кг." + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).FontSize(10)
                    .Append(proforma[24]).FontSize(10).Append(proforma[25]).FontSize(10).UnderlineStyle(UnderlineStyle.singleLine).SetLineSpacing(LineSpacingType.Line, 1.7f);

                Paragraph p13 = document.InsertParagraph();
                p13.Append(proforma[26] + " ").FontSize(10).Append("ПРР: " + lineInfTextBox.Text + ", ГРН" + Environment.NewLine).FontSize(10).UnderlineStyle(UnderlineStyle.singleLine)
                    .Append("                            ").FontSize(10);

                if (arena.Checked == true)
                {
                    p9.Append(proforma[27]).FontSize(10).UnderlineStyle(UnderlineStyle.singleLine);
                }
                else
                {
                    p9.Append(proforma[31]).FontSize(10).UnderlineStyle(UnderlineStyle.singleLine);
                }

                foreach (var item in document.Paragraphs)
                {
                    item.Font("Times New Roman").Bold();
                }
                secondPage = false;
            }
            try
            {
                document.Save();
                Process.Start(@"" + Properties.Settings.Default["SavingPath"].ToString() + "\\" + linesComboBox.Text + "\\" + vesselsComboBox.Text + "\\" + senderComboBox.Text + " - Прч. " + errendNumber.Text + " - " + vr.Values.Count.ToString() + " конт. - " + countryСomboBox.Text + ".docx");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void FillCombatBox()
        {
            try
            {
                ds = new DataSet();
                using (SqlCeConnection conn = new SqlCeConnection(sqlConnection))
                {
                    conn.Open();

                    string queryLines = "select Name, Payer from Lines";
                    SqlCeDataAdapter da = new SqlCeDataAdapter(queryLines, conn);
                    da.Fill(ds, "Lines");
                    linesComboBox.DisplayMember = "Name";
                    linesComboBox.ValueMember = "Name";
                    linesComboBox.DataSource = ds.Tables["Lines"];
                    linesComboBox.SelectedIndex = -1;

                    string querySeller = "select ShortName, Name, Address, Cod from Sender";
                    da = new SqlCeDataAdapter(querySeller, conn);
                    da.Fill(ds, "Sender");
                    senderComboBox.DisplayMember = "ShortName";
                    senderComboBox.ValueMember = "ShortName";
                    senderComboBox.DataSource = ds.Tables["Sender"];
                    senderComboBox.SelectedIndex = -1;

                    string queryReceiver = "select Name, Address, Country, Sender from Receiver";
                    da = new SqlCeDataAdapter(queryReceiver, conn);
                    da.Fill(ds, "Receiver");
                    receiverComboBox.DisplayMember = "Name";
                    receiverComboBox.ValueMember = "Name";
                    receiverComboBox.DataSource = new DataView(ds.Tables["Receiver"]).ToTable(true, "Name");
                    receiverComboBox.SelectedIndex = -1;

                    string queryFlag = "select Name, Flag from Vessels";
                    da = new SqlCeDataAdapter(queryFlag, conn);
                    da.Fill(ds, "Vessels");

                    string queryPortT = "SELECT Name from PortT";
                    da = new SqlCeDataAdapter(queryPortT, conn);
                    da.Fill(ds, "PortT");
                    potComboBox.DisplayMember = "Name";
                    potComboBox.ValueMember = "Name";
                    potComboBox.DataSource = ds.Tables["PortT"];
                    potComboBox.SelectedIndex = -1;

                    string queryCountry = "SELECT Name, Port from CountryPort";
                    da = new SqlCeDataAdapter(queryCountry, conn);
                    da.Fill(ds, "CountryPort");
                    countryСomboBox.DisplayMember = "Name";
                    countryСomboBox.ValueMember = "Name";
                    countryСomboBox.DataSource = new DataView(ds.Tables["CountryPort"]).ToTable(true, "Name");
                    countryСomboBox.SelectedIndex = -1;
                    podComboBox.DisplayMember = "Port";
                    podComboBox.ValueMember = "Port";
                    podComboBox.DataSource = ds.Tables["CountryPort"];
                    podComboBox.SelectedIndex = -1;

                    string queryCargo = "SELECT Name, Cod, ShortName from Cargo";
                    da = new SqlCeDataAdapter(queryCargo, conn);
                    da.Fill(ds, "Cargo");
                    cargoNameСomboBox.DisplayMember = "Name";
                    cargoNameСomboBox.ValueMember = "Name";
                    cargoNameСomboBox.DataSource = ds.Tables["Cargo"];
                    cargoNameСomboBox.SelectedIndex = -1;
                    cargoCodComboBox.DisplayMember = "Cod";
                    cargoCodComboBox.ValueMember = "Cod";
                    cargoCodComboBox.DataSource = ds.Tables["Cargo"];
                    File.Copy(dbPath, @"db.sdf", true);
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occured! " + ex);
            }
            cargoCodComboBox.SelectedIndex = cargoNameСomboBox.SelectedIndex;
        }


        private void linesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (linesComboBox.SelectedIndex != -1)
            {
                lineInfTextBox.Text = ds.Tables["Lines"].Rows[linesComboBox.SelectedIndex]["Payer"].ToString();
                lineInfTextBox.Text = ds.Tables["Lines"].Select("Name = '" + linesComboBox.Text + "'")[0]["Payer"].ToString();
            }
            else
            {
                lineInfTextBox.Text = "";
            }
        }

        private void senderComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (senderComboBox.SelectedIndex != -1)
            {
                senderNameTextBox.Text = ds.Tables["Sender"].Rows[senderComboBox.SelectedIndex]["Name"].ToString();
                senderAddressTextBox.Text = ds.Tables["Sender"].Rows[senderComboBox.SelectedIndex]["Address"].ToString();
                senderCodTextBox.Text = ds.Tables["Sender"].Rows[senderComboBox.SelectedIndex]["Cod"].ToString();
            }
            else
            {
                senderNameTextBox.Text = "";
                senderAddressTextBox.Text = "";
                senderCodTextBox.Text = "";
            }
        }

        private void receiverComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (receiverComboBox.SelectedIndex != -1)
            {
                receiverAddressTextBox.Text = ds.Tables["Receiver"].Select("Name = '" + receiverComboBox.Text + "'")[0]["Address"].ToString();
            }
            else
            {
                receiverComboBox.Text = "";
                receiverAddressTextBox.Text = "";
            }
        }

        private void vesselsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            voyageComboBox.Items.Clear();
            for (int i = 0; i < vessels.Length; i++)
            {
                if (vesselsComboBox.SelectedIndex != -1 && vessels[i][2] == vesselsComboBox.SelectedItem.ToString())
                {
                    voyageComboBox.Items.Add(vessels[i][6]);
                    voyageComboBox.SelectedIndex = 0;
                }

            }
            if (vesselsComboBox.SelectedIndex == -1)
            {
                voyageComboBox.Text = "";
                seTextBox.Text = "";
                flagTextBox.Text = "";
            }
            if (ds.Tables["Vessels"].Select("Name = '" + vesselsComboBox.Text + "'").Length != 0)
            {
                flagTextBox.Text = ds.Tables["Vessels"].Select("Name = '" + vesselsComboBox.Text + "'")[0]["Flag"].ToString();
            }
            else
            {
                flagTextBox.Text = "";
            }
        }

        private void voyageComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < vessels.Length; i++)
            {
                if (vessels[i][6] == voyageComboBox.SelectedItem.ToString())
                {
                    seTextBox.Text = vessels[i][1];
                }
            }
        }

        private void countryСomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (countryСomboBox.SelectedIndex != -1)
            {
                podComboBox.DisplayMember = "Port";
                podComboBox.ValueMember = "Port";
                podComboBox.DataSource = ds.Tables["CountryPort"].Select("Name = '" + countryСomboBox.SelectedValue.ToString() + "'").CopyToDataTable();
            }
            else
            {

                podComboBox.DataSource = ds.Tables["CountryPort"];
                podComboBox.SelectedIndex = -1;
            }
        }

        private void cargoNameСomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cargoNameСomboBox.SelectedIndex > 0)
            {
                cargoCodComboBox.SelectedIndex = cargoNameСomboBox.SelectedIndex;
            }
        }

        private void detectSomeData_Click(object sender, EventArgs e)
        {
            ClearFields();
            vr = gXLS.ReadDataFromXMLErrend(textBox1.Text);
            if (vr != null)
            {
                errendNumber.Text = vr.Values[0][2].ToString();
                linesComboBox.Text = vr.Values[0][1].ToString();
                senderComboBox.Text = vr.Values[0][5].ToString();
                vesselsComboBox.Text = vr.Values[0][8].ToString();
                countryСomboBox.Text = vr.Values[0][7].ToString();
                bookingTextBox.Text = vr.Values[0][3].ToString();
                try
                {
                    cargoNameСomboBox.SelectedValue = ds.Tables["Cargo"].Select("ShortName = '" + vr.Values[0][6].ToString() + "'")[0]["Name"].ToString();
                }
                catch (Exception)
                {

                }
                try
                {
                    receiverComboBox.DataSource = ds.Tables["Receiver"].Select("Country = '" + countryСomboBox.Text + "' AND Sender = '" + senderComboBox.Text + "'").CopyToDataTable();
                }
                catch (Exception)
                {

                }
            }
        }

        private void firstPoint_TextChanged(object sender, EventArgs e)
        {
            if (firstPoint.Text != "" && vr != null)
            {
                secondPont.Text = (Convert.ToInt32(firstPoint.Text) + Convert.ToInt32(vr.Values[0][0]) - 1).ToString();
            }
        }

        private void addReceiver_Click(object sender, EventArgs e)
        {
            if (ds.Tables["Receiver"].Select("Name = '" + receiverComboBox.Text.ToUpper() + "' AND Country = '" + countryСomboBox.Text.ToUpper() + "' AND Sender = '" + senderComboBox.Text.ToUpper() + "'").Length == 0)
            {
                CheckFields();
                if (!invalidField)
                {
                    ChangeFieldsColor();
                    using (SqlCeConnection conn = new SqlCeConnection(sqlConnection))
                    {
                        conn.Open();
                        SqlCeCommand cmd = new SqlCeCommand("INSERT INTO Receiver(Name, Address, Country, Sender) VALUES (@Name, @Address, @Country, @Sender)", conn);
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlCeParameter("Name", receiverComboBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("Address", receiverAddressTextBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("Country", countryСomboBox.Text));
                        cmd.Parameters.Add(new SqlCeParameter("Sender", senderComboBox.Text));
                        cmd.ExecuteNonQuery();
                        ds.Tables["Receiver"].Clear();
                        string queryReceiver = "select Name, Address, Country, Sender from Receiver";
                        SqlCeDataAdapter da = new SqlCeDataAdapter(queryReceiver, conn);
                        da.Fill(ds, "Receiver");
                        conn.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Получатель уже существует!");
            }
        }

        private void addSender_Click(object sender, EventArgs e)
        {
            if (ds.Tables["Sender"].Select("Name = '" + senderNameTextBox.Text.ToUpper() + "'").Length == 0)
            {
                CheckFields();
                if (!invalidField)
                {
                    ChangeFieldsColor();
                    using (SqlCeConnection conn = new SqlCeConnection(sqlConnection))
                    {
                        conn.Open();
                        SqlCeCommand cmd = new SqlCeCommand("INSERT INTO Sender(ShortName, Name, Address, Cod) VALUES (@ShortName, @Name, @Address, @Cod)", conn);
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlCeParameter("ShortName", senderComboBox.Text));
                        cmd.Parameters.Add(new SqlCeParameter("Name", senderNameTextBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("Address", senderAddressTextBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("Cod", senderCodTextBox.Text));
                        cmd.ExecuteNonQuery();
                        ds.Tables["Sender"].Clear();
                        string querySeller = "select ShortName, Name, Address, Cod from Sender";
                        SqlCeDataAdapter da = new SqlCeDataAdapter(querySeller, conn);
                        da.Fill(ds, "Sender");
                        conn.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Отправитель уже существует!");
            }

        }

        private void addCargo_Click(object sender, EventArgs e)
        {
            if (ds.Tables["Cargo"].Select("Name = '" + cargoNameСomboBox.Text.ToUpper() + "' AND Cod = '" + cargoCodComboBox.Text + "'").Length == 0)
            {
                CheckFields();
                if (!invalidField)
                {
                    ChangeFieldsColor();
                    if (cargoShortName.Text != "")
                    {
                        cargoShortName.BackColor = System.Drawing.Color.White;
                        using (SqlCeConnection conn = new SqlCeConnection(sqlConnection))
                        {
                            conn.Open();
                            SqlCeCommand cmd = new SqlCeCommand("INSERT INTO Cargo(Name, Cod, ShortName) VALUES (@Name, @Cod, @ShortName)", conn);
                            cmd.Parameters.Clear();
                            cmd.Parameters.Add(new SqlCeParameter("Name", cargoNameСomboBox.Text.ToUpper()));
                            cmd.Parameters.Add(new SqlCeParameter("Cod", cargoCodComboBox.Text));
                            cmd.Parameters.Add(new SqlCeParameter("ShortName", cargoShortName.Text.ToUpper()));
                            cmd.ExecuteNonQuery();
                            ds.Tables["Cargo"].Clear();
                            string queryCargo = "SELECT Name, Cod, ShortName from Cargo";
                            SqlCeDataAdapter da = new SqlCeDataAdapter(queryCargo, conn);
                            da.Fill(ds, "Cargo");
                            conn.Close();
                        }
                    }
                    else
                    {
                        cargoShortName.BackColor = System.Drawing.Color.LightCoral;
                    }
                }
            }
            else
            {
                MessageBox.Show("Груз уже существует!");
            }
        }

        private void addPortD_Click(object sender, EventArgs e)
        {
            if (ds.Tables["CountryPort"].Select("Port = '" + podComboBox.Text.ToUpper() + "'").Length == 0)
            {
                CheckFields();
                if (!invalidField)
                {
                    ChangeFieldsColor();
                    using (SqlCeConnection conn = new SqlCeConnection(sqlConnection))
                    {
                        conn.Open();
                        SqlCeCommand cmd = new SqlCeCommand("INSERT INTO CountryPort(Name, Port) VALUES (@Name, @Port)", conn);
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlCeParameter("Name", countryСomboBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("Port", podComboBox.Text.ToUpper()));
                        cmd.ExecuteNonQuery();
                        ds.Tables["CountryPort"].Clear();
                        string queryCountry = "SELECT Name, Port from CountryPort";
                        SqlCeDataAdapter da = new SqlCeDataAdapter(queryCountry, conn);
                        da.Fill(ds, "CountryPort");
                        conn.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Порт уже существует!");
            }
        }

        private void addPortT_Click(object sender, EventArgs e)
        {
            if (ds.Tables["PortT"].Select("Name = '" + potComboBox.Text.ToUpper() + "'").Length == 0)
            {
                CheckFields();
                if (!invalidField)
                {
                    ChangeFieldsColor();
                    using (SqlCeConnection conn = new SqlCeConnection(sqlConnection))
                    {
                        conn.Open();
                        SqlCeCommand cmd = new SqlCeCommand("INSERT INTO PortT(Name) VALUES (@Name)", conn);
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlCeParameter("Name", potComboBox.Text.ToUpper()));
                        cmd.ExecuteNonQuery();
                        ds.Tables["PortT"].Clear();
                        string queryPortT = "SELECT Name from PortT";
                        SqlCeDataAdapter da = new SqlCeDataAdapter(queryPortT, conn);
                        da.Fill(ds, "PortT");
                        conn.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Порт уже существует!");
            }
        }

        private void addVessel_Click(object sender, EventArgs e)
        {
            if (ds.Tables["Vessels"].Select("Name = '" + vesselsComboBox.Text.ToUpper() + "'").Length == 0)
            {
                CheckFields();
                if (!invalidField)
                {
                    ChangeFieldsColor();
                    using (SqlCeConnection conn = new SqlCeConnection(sqlConnection))
                    {
                        conn.Open();
                        SqlCeCommand cmd = new SqlCeCommand("INSERT INTO Vessels(Name, Flag) VALUES (@Name, @Flag)", conn);
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlCeParameter("Name", vesselsComboBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("Flag", flagTextBox.Text.ToUpper()));
                        cmd.ExecuteNonQuery();
                        ds.Tables["Vessels"].Clear();
                        string queryFlag = "select Name, Flag from Vessels";
                        SqlCeDataAdapter da = new SqlCeDataAdapter(queryFlag, conn);
                        da.Fill(ds, "Vessels");
                        conn.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Судно уже существует!");
            }
        }

        private void clFields_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            ClearFields();
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form f = new Form2();
            if (f.ShowDialog() == DialogResult.Cancel)
            {
                using (Process myProcess = new Process())
                {
                    Application.Exit();
                    Process.Start("Errend.exe", "");
                }
            }
        }


        private void базаДанныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form f = new Form3();
            if (f.ShowDialog() == DialogResult.Cancel)
            {
                using (Process myProcess = new Process())
                {
                    Application.Exit();
                    Process.Start("Errend.exe", "");
                }
            }
        }

        private void bKPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
            Thread tr = new Thread(OpenForm);
            tr.SetApartmentState(ApartmentState.STA);
            tr.Start();
        }

        private void OpenForm()
        {
            Application.Run(new Form4());
        }

        private void arena_CheckedChanged(object sender, EventArgs e)
        {
            if(arena.Checked == true)
            {
                ugl.Checked = false;
            }
        }

        private void ugl_CheckedChanged(object sender, EventArgs e)
        {
            if (ugl.Checked == true)
            {
                arena.Checked = false;
            }
        }
    }
}
