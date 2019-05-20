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
using Errend.Core.BKP;
using System.Globalization;
using System.Data.SqlServerCe;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace Errend
{
    public partial class Form4 : Form
    {
        string[] proforma = File.ReadAllLines(@"SomeTXTBKP.txt", Encoding.GetEncoding(1251));
        string[] tableCMA = File.ReadAllLines(@"TableCMA.txt", Encoding.GetEncoding(1251));
        private double weight;
        string[] months = new string[] { "января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря" };
        GoogleXLS gXML = new GoogleXLS();
        DataSet ds;
        GoogleXLS gXLS = new GoogleXLS();
        string dbPath = Properties.Settings.Default["DBPath"].ToString();
        string sqlConnection = @"Data Source = " + Properties.Settings.Default["DBPath"].ToString() + "; Persist Security Info=False";
        ParserWorker<string[]> parserBKP;
        string[,] vesselsBKP;
        ValueRange vr = null;
        bool invalidField = false;

        public Form4()
        {
            InitializeComponent();
            string commonFolder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            string myAppFolder = Path.Combine(commonFolder, "Errend");
            AppDomain.CurrentDomain.SetData("DataDirectory", myAppFolder);
            parserBKP = new ParserWorker<string[]>(new BkpParser());
            parserBKP.OnNewData += Parser_OnNewDataBKP;
            parserBKP.Settings = new BkpSettings();
            parserBKP.Start();
            FillCombatBox();
            ContData();
        }

        private void Parser_OnNewDataBKP(object arg1, string[] arg2)
        {
            vesselsBKP = new string[arg2.Length, 2];
            for (int i = 0; i < arg2.Length; i += 2)
            {
                vesselsBKP[i, 0] = arg2[i].Trim();
                vesselsBKP[i, 1] = arg2[i + 1].Trim();
            }

            for (int i = 0; i < vesselsBKP.Length / 2; i++)
            {
                if (vesselsBKP[i,0] != null && !vesselsComboBox.Items.Contains(vesselsBKP[i,0]))
                {
                    vesselsComboBox.Items.Add(vesselsBKP[i,0]);
                }

            }
        }

        private void ContData()
        {
            string[] contSize = new string[] { "20", "40", "45" };
            string[] contType = new string[] { "DV", "HC", "HP", "RF", "PL" };
            contSizeComboBox.Items.AddRange(contSize);
            contTypeComboBox.Items.AddRange(contType);
        }

        private void CreateDocBKP(ValueRange vr)
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
            if (vr.Values.Count > 15)
            {
                secondPage = true;
                amountContPage1 = 15;
                amounrContPage2 = vr.Values.Count - 15;
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
            Xceed.Words.NET.Picture p = img.CreatePicture();
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
            Table tFooter = footer.InsertTable(2, 6);
            tFooter.Rows[1].Height = 50;
            tFooter.Rows[0].Cells[0].Paragraphs.First().IndentationBefore = -1.5f;
            tFooter.Rows[0].Cells[0].MarginLeft = 50;
            tFooter.Rows[0].Cells[3].MarginRight = 10;
            tFooter.Rows[0].Cells[0].Width = 160;
            tFooter.Rows[0].Cells[1].Width = 50;
            tFooter.Rows[0].Cells[2].Width = 50;
            tFooter.Rows[0].Cells[3].Width = 160;
            tFooter.Rows[0].Cells[4].Width = 160;
            tFooter.Rows[0].Cells[5].Width = 160;
            tFooter.Rows[0].Cells[0].Paragraphs.First().Append("Экспедитор" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Responsible person and company stamp").Alignment = Alignment.center;
            tFooter.Rows[0].Cells[1].Paragraphs.First().Append("Ветконтроль" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Veterinary inspection").Alignment = Alignment.center;
            tFooter.Rows[0].Cells[2].Paragraphs.First().Append("Карантин" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Quarantine inspection").Alignment = Alignment.center;
            tFooter.Rows[0].Cells[3].Paragraphs.First().Append("Таможня" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Customs stamp").Alignment = Alignment.center;
            tFooter.Rows[0].Cells[4].Paragraphs.First().Append("Линейный агент" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Line agent stamp").Alignment = Alignment.center;
            tFooter.Rows[0].Cells[5].Paragraphs.First().Append("Судовой агент" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Vessel’s agent stamp").Alignment = Alignment.center;
            tFooter.Rows[1].Cells[0].Paragraphs[0].Append("Глоба В. Л.\n").Alignment = Alignment.left;
            tFooter.Rows[1].Cells[0].Paragraphs[0].Append("Тел.050-341-89-12").Alignment = Alignment.left;
            tFooter.Rows[1].Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            foreach (var item in tFooter.Paragraphs)
            {
                item.Font("Times New Roman").FontSize(9).Bold();
            }
            document.MarginFooter = 15;

            Paragraph p1 = document.InsertParagraph();
            p1.Append(proforma[0] + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(proforma[1])
                .Append(proforma[2] + "«" + linesComboBox.Text.ToUpper() + "»" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(arena.Checked == true ? proforma[3] : proforma[48])
                .Append(proforma[4] + "«" + linesComboBox.Text.ToUpper() + "»").SpacingAfter(8);
            

            DateTime date = DateTime.Now;
            Paragraph p2 = document.InsertParagraph();
            p2.Append(proforma[5] + " " + errendNumber.Text + " от(dd) «" + date.Day.ToString("d2") + "» " + months[date.Month - 1] + " " + date.Year + " г.(year)" + Environment.NewLine)
                .Append(proforma[6]).Append(proforma[7]).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[8] + Environment.NewLine)
                .Append(proforma[9]).Append(proforma[10]).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[11]).SpacingAfter(10).Alignment = Alignment.center;

            Paragraph p3 = document.InsertParagraph();
            p3.Append(proforma[12] + " " + senderNameRusTextBox.Text.ToUpper() + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(proforma[13] + " " + senderNameEngTextBox.Text.ToUpper() + Environment.NewLine)
                .Append(senderAddressRusTextBox.Text.ToUpper() + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(senderAddressEngTextBox.Text.ToUpper() + Environment.NewLine)
                .Append("Код: " + senderCodTextBox.Text).SpacingAfter(8);

            Paragraph p4 = document.InsertParagraph();
            p4.Append(proforma[14] + " " + receiverComboBox.Text.ToUpper() + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
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
            t1.Rows[1].Cells[0].Paragraphs.First().Append(proforma[15] + " «" + vesselsComboBox.Text.ToUpper() + "»" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[16] + " «" + vesselsComboBox.Text.ToUpper() + "»").SpacingAfter(8);
            t1.Rows[1].Cells[1].Paragraphs.First().Append(proforma[17] + " " + voyageTextBox.Text.ToUpper() + "  " + proforma[18]  + " " + flagRusTextBox.Text.ToUpper() + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[19] + " " + voyageTextBox.Text.ToUpper() + "  " + proforma[20] + " " + flagEngTextBox.Text.ToUpper()).SpacingAfter(8);
            t1.Rows[2].Cells[0].Paragraphs.First().Append(proforma[21] + " " + proforma[22] + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[23] + " " + proforma[24]).SpacingAfter(8);
            t1.Rows[2].Cells[1].Paragraphs.First().Append(proforma[25] + " " + podRusTextBox.Text.ToUpper() + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[26] + " " + podComboBox.Text.ToUpper()).SpacingAfter(8);
            t1.Rows[3].Cells[0].Paragraphs.First().Append(proforma[27] + " " + potRusTextBox.Text.ToUpper() + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[28] + " " + potComboBox.Text.ToUpper()).SpacingAfter(8);
            t1.Rows[4].Cells[0].Paragraphs.First().Append(proforma[29] + " " + cargoNameСomboBox.Text.ToUpper() + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[30] + " " + cargoNameEngTextBox.Text.ToUpper()).SpacingAfter(8);
            t1.Rows[4].Cells[1].Paragraphs.First().Append(proforma[31] + " " + cargoCodComboBox.Text + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[32] + " " + cargoCodComboBox.Text);
            document.InsertTable(t1);

            Paragraph p5 = document.InsertParagraph();
            p5.Append(proforma[33] + " " + bookingTextBox.Text.ToUpper() + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append(proforma[34] + " " + bookingTextBox.Text.ToUpper());

            Table t2 = document.AddTable(amountContPage1 + 1, 8);
            t2.Rows[0].Cells[0].MarginLeft = 40;
            t2.Rows[0].Cells[0].Paragraphs.First().Append("Номер контейнера" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Container number").Alignment = Alignment.center;
            t2.Rows[0].Cells[0].Paragraphs.First().IndentationBefore = -1f;
            t2.Rows[0].Cells[1].Paragraphs.First().Append("Тип" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Type").Alignment = Alignment.center;
            t2.Rows[0].Cells[2].Paragraphs.First().Append("Кол-во Мест" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Number of units").Alignment = Alignment.center;
            t2.Rows[0].Cells[3].Paragraphs.First().Append("Вес груза нетто" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Cargo net weight").Alignment = Alignment.center;
            t2.Rows[0].Cells[4].Paragraphs.First().Append("Вес груза брутто" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Cargo gross weight").Alignment = Alignment.center;
            t2.Rows[0].Cells[5].Paragraphs.First().Append("Подтвер-жденный вес к-ра брутто" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("VGM").Alignment = Alignment.center;
            t2.Rows[0].Cells[6].Paragraphs.First().Append("Пломбы" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Seals").Alignment = Alignment.center;
            t2.Rows[0].Cells[6].Width = 250;
            t2.Rows[0].Cells[7].MarginRight = 10;
            t2.Rows[0].Cells[7].Paragraphs.First().Append("ГТД" + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).Append("Customs declaration").Alignment = Alignment.center;
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

            Paragraph p6 = document.InsertParagraph();
            p6.Append(proforma[35] + " " + vr.Values.Count + "x" + contSizeComboBox.Text + " контейнер(ов). ВЕС \u2013" + " " + strWeight + " кг." + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(proforma[36] + " " + vr.Values.Count + "x" + contSizeComboBox.Text + " container(s). ВЕС \u2013" + " " + strWeight + " kg.").SpacingBefore(5).SpacingAfter(8);

            Paragraph p7 = document.InsertParagraph();
            p7.Append(proforma[37] + " " + proforma[38] + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(proforma[39] + " " + proforma[40]).SpacingAfter(8);

            Paragraph p8 = document.InsertParagraph();
            p8.Append(proforma[41] + " ПРР " + lineInfRusTextBox.Text.ToUpper() + ", ГРН." +  Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(proforma[42] + " THC " + lineInfEngTextBox.Text.ToUpper() + ", UAH.").SpacingAfter(8);

            Paragraph p9 = document.InsertParagraph();
            p9.Append(proforma[43] + " " + (arena.Checked == true ? proforma[44] : proforma[49]) + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine)
                .Append(proforma[45] + " " + (arena.Checked == true ? proforma[46] : proforma[50]));

            foreach (var item in document.Paragraphs)
            {
                item.Font("Times New Roman").FontSize(9).Bold();
            }
            document.Paragraphs[1].FontSize(11);

            if (secondPage)
            {
                p9.InsertPageBreakAfterSelf();
                Paragraph p10 = document.InsertParagraph();
                p10.Append(proforma[47] + Environment.NewLine + " " + errendNumber.Text + " от(dd) «" + date.Day.ToString("d2") + "» " + months[date.Month - 1] + " " + date.Year + " г.(year)").SpacingAfter(8).Alignment = Alignment.center;

                Table t3 = document.AddTable(amounrContPage2 + 1, 8);
                t3.Rows[0].Cells[0].MarginLeft = 40;
                t3.Rows[0].Cells[0].Paragraphs.First().Append("Номер контейнера" + Environment.NewLine).FontSize(9).UnderlineStyle(UnderlineStyle.singleLine).Append("Container number").FontSize(9).Alignment = Alignment.center;
                t3.Rows[0].Cells[0].Paragraphs.First().IndentationBefore = -1f;
                t3.Rows[0].Cells[1].Paragraphs.First().Append("Тип" + Environment.NewLine).FontSize(9).UnderlineStyle(UnderlineStyle.singleLine).Append("Type").FontSize(9).Alignment = Alignment.center;
                t3.Rows[0].Cells[2].Paragraphs.First().Append("Кол-во Мест" + Environment.NewLine).FontSize(9).UnderlineStyle(UnderlineStyle.singleLine).Append("Number of units").FontSize(9).Alignment = Alignment.center;
                t3.Rows[0].Cells[3].Paragraphs.First().Append("Вес груза нетто" + Environment.NewLine).FontSize(9).UnderlineStyle(UnderlineStyle.singleLine).Append("Cargo net weight").FontSize(9).Alignment = Alignment.center;
                t3.Rows[0].Cells[4].Paragraphs.First().Append("Вес груза брутто" + Environment.NewLine).FontSize(9).UnderlineStyle(UnderlineStyle.singleLine).Append("Cargo gross weight").FontSize(9).Alignment = Alignment.center;
                t3.Rows[0].Cells[5].Paragraphs.First().Append("Подтвер-жденный вес к-ра брутто" + Environment.NewLine).FontSize(9).UnderlineStyle(UnderlineStyle.singleLine).Append("VGM").FontSize(9).Alignment = Alignment.center;
                t3.Rows[0].Cells[6].Paragraphs.First().Append("Пломбы" + Environment.NewLine).FontSize(9).UnderlineStyle(UnderlineStyle.singleLine).Append("Seals").FontSize(9).Alignment = Alignment.center;
                t3.Rows[0].Cells[6].Width = 250;
                t3.Rows[0].Cells[7].MarginRight = 10;
                t3.Rows[0].Cells[7].Paragraphs.First().Append("ГТД" + Environment.NewLine).FontSize(9).UnderlineStyle(UnderlineStyle.singleLine).Append("Customs declaration").FontSize(9).Alignment = Alignment.center;
                for (int i = 0; i < amounrContPage2; i++)
                {
                    t3.Rows[i + 1].Cells[0].Paragraphs.First().Append(vr.Values[i + 15][0].ToString()).FontSize(9).Alignment = Alignment.center;
                    t3.Rows[i + 1].Cells[1].Paragraphs.First().Append(contSizeComboBox.Text + contTypeComboBox.Text).FontSize(9).Alignment = Alignment.center;
                    t3.Rows[i + 1].Cells[2].Paragraphs.First().Append(vr.Values[i + 15][13].ToString().ToUpper() == "НАВАЛ" || vr.Values[i + 15][13].ToString().ToUpper() == "" ? "НАВАЛ" : vr.Values[i + 15][13].ToString()).FontSize(9).Alignment = Alignment.center;
                    t3.Rows[i + 1].Cells[3].Paragraphs.First().Append(vr.Values[i + 15][4].ToString()).FontSize(9).Alignment = Alignment.center;
                    t3.Rows[i + 1].Cells[4].Paragraphs.First().Append(vr.Values[i + 15][5].ToString()).FontSize(9).Alignment = Alignment.center;
                    t3.Rows[i + 1].Cells[5].Paragraphs.First().Append(vr.Values[i + 15][6].ToString()).FontSize(9).Alignment = Alignment.center;
                    if (vr.Values[i].Count == 20)
                    {
                        t3.Rows[i + 1].Cells[6].Paragraphs.First().Append(vr.Values[i + 15][19].ToString()).FontSize(9).Alignment = Alignment.center;
                    }
                    else
                    {
                        t3.Rows[i + 1].Cells[6].Paragraphs.First().Append("").FontSize(9).Alignment = Alignment.center;
                    }
                    t3.Rows[i + 1].Cells[7].Paragraphs.First().Append(vr.Values[i + 15][2].ToString()).FontSize(9).Alignment = Alignment.center;
                    if (vr.Values[i + 15][2].ToString() == "")
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
                p11.Append(proforma[35] + " " + vr.Values.Count + "x" + contSizeComboBox.Text + " контейнер(ов). ВЕС \u2013" + " " + strWeight + " кг." + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).FontSize(9)
                    .Append(proforma[36] + " " + vr.Values.Count + "x" + contSizeComboBox.Text + " container(s). ВЕС \u2013" + " " + strWeight + " kg.").SpacingBefore(5).SpacingAfter(8).FontSize(9);

                Paragraph p12 = document.InsertParagraph();
                p12.Append(proforma[37] + " " + proforma[38] + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).FontSize(9)
                    .Append(proforma[39] + " " + proforma[40]).SpacingAfter(8).FontSize(9);

                Paragraph p13 = document.InsertParagraph();
                p13.Append(proforma[41] + " ПРР " + lineInfRusTextBox.Text.ToUpper() + ", ГРН." + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).FontSize(9)
                    .Append(proforma[42] + " THC " + lineInfEngTextBox.Text.ToUpper() + ", UAH.").SpacingAfter(8).FontSize(9);

                Paragraph p14 = document.InsertParagraph();
                p14.Append(proforma[43] + " " + (arena.Checked == true ? proforma[44] : proforma[49]) + Environment.NewLine).UnderlineStyle(UnderlineStyle.singleLine).FontSize(9)
                    .Append(proforma[45] + " " + (arena.Checked == true ? proforma[46] : proforma[50])).FontSize(9);

                foreach (var item in document.Paragraphs)
                {
                    item.Font("Times New Roman").Bold();
                }
                secondPage = false;
            }
            if (linesComboBox.Text == "CMA-CGM")
            {
                CreateXLS(vr, voyageTextBox.Text, vesselsComboBox.Text, errendNumber.Text, "Arena Marine", "--LEAVE BLANK--", senderNameEngTextBox.Text + "\n" + senderAddressEngTextBox.Text, receiverComboBox.Text + "\n" + receiverAddressTextBox.Text, podComboBox.Text, cargoNameEngTextBox.Text);
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

                    string queryLines = "select * from Lines";
                    SqlCeDataAdapter da = new SqlCeDataAdapter(queryLines, conn);
                    da.Fill(ds, "Lines");
                    linesComboBox.DisplayMember = "Name";
                    linesComboBox.ValueMember = "Name";
                    linesComboBox.DataSource = ds.Tables["Lines"];
                    linesComboBox.SelectedIndex = -1;

                    string querySeller = "select * from Sender";
                    da = new SqlCeDataAdapter(querySeller, conn);
                    da.Fill(ds, "Sender");
                    senderComboBox.DisplayMember = "ShortName";
                    senderComboBox.ValueMember = "ShortName";
                    senderComboBox.DataSource = ds.Tables["Sender"];
                    senderComboBox.SelectedIndex = -1;

                    string queryReceiver = "select * from Receiver";
                    da = new SqlCeDataAdapter(queryReceiver, conn);
                    da.Fill(ds, "Receiver");
                    receiverComboBox.DisplayMember = "Name";
                    receiverComboBox.ValueMember = "Name";
                    receiverComboBox.DataSource = new DataView(ds.Tables["Receiver"]).ToTable(true, "Name");
                    receiverComboBox.SelectedIndex = -1;

                    string queryFlag = "select * from Vessels";
                    da = new SqlCeDataAdapter(queryFlag, conn);
                    da.Fill(ds, "Vessels");

                    string queryPortT = "SELECT * from PortT";
                    da = new SqlCeDataAdapter(queryPortT, conn);
                    da.Fill(ds, "PortT");
                    potComboBox.DisplayMember = "Name";
                    potComboBox.ValueMember = "Name";
                    potComboBox.DataSource = ds.Tables["PortT"];
                    potComboBox.SelectedIndex = -1;

                    string queryCountry = "SELECT * from CountryPort";
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

                    string queryCargo = "SELECT * from Cargo";
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

        private void CreateDocument_Click(object sender, EventArgs e)
        {
            CheckFields();
            if (invalidField)
            {
                invalidField = false;
                return;
            }
            ChangeFieldsColor();
            CreateDocBKP(gXML.ReadDataFromGoogleXML(firstPoint.Text, secondPont.Text));
        }

        private void linesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (linesComboBox.SelectedIndex != -1)
            {
                lineInfRusTextBox.Text = ds.Tables["Lines"].Select("Name = '" + linesComboBox.Text + "'")[0]["Payer"].ToString();
                lineInfEngTextBox.Text = ds.Tables["Lines"].Select("Name = '" + linesComboBox.Text + "'")[0]["PayerENG"].ToString();
            }
            else
            {
                lineInfRusTextBox.Text = "";
                lineInfEngTextBox.Text = "";
            }
        }

        private void senderComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (senderComboBox.SelectedIndex != -1)
            {
                senderNameRusTextBox.Text = ds.Tables["Sender"].Rows[senderComboBox.SelectedIndex]["Name"].ToString();
                senderAddressRusTextBox.Text = ds.Tables["Sender"].Rows[senderComboBox.SelectedIndex]["Address"].ToString();
                senderCodTextBox.Text = ds.Tables["Sender"].Rows[senderComboBox.SelectedIndex]["Cod"].ToString();
                senderNameEngTextBox.Text = ds.Tables["Sender"].Rows[senderComboBox.SelectedIndex]["NameENG"].ToString();
                senderAddressEngTextBox.Text = ds.Tables["Sender"].Rows[senderComboBox.SelectedIndex]["AddressENG"].ToString();
            }
            else
            {
                senderComboBox.Text = "";
                senderNameRusTextBox.Text = "";
                senderAddressRusTextBox.Text = "";
                senderNameEngTextBox.Text = "";
                senderAddressEngTextBox.Text = "";
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
            for (int i = 0; i < vesselsBKP.Length / 2; i++)
            {
                if (vesselsComboBox.SelectedIndex != -1 && vesselsBKP[i,0] == vesselsComboBox.SelectedItem.ToString() && vesselsBKP[i, 0] != null)
                {
                    voyageTextBox.Text = vesselsBKP[i,1];
                }
            }
            if (vesselsComboBox.SelectedIndex == -1)
            {
                voyageTextBox.Text = "";
                flagEngTextBox.Text = "";
                flagRusTextBox.Text = "";
            }
            if (ds.Tables["Vessels"].Select("Name = '" + vesselsComboBox.Text + "'").Length != 0)
            {
                flagEngTextBox.Text = ds.Tables["Vessels"].Select("Name = '" + vesselsComboBox.Text + "'")[0]["Flag"].ToString();
                flagRusTextBox.Text = ds.Tables["Vessels"].Select("Name = '" + vesselsComboBox.Text + "'")[0]["FlagRus"].ToString();
            }
            else
            {
                flagEngTextBox.Text = "";
                flagRusTextBox.Text = "";
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
                countryСomboBox.Text = "";
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

            if (cargoNameСomboBox.SelectedIndex != -1)
            {
                cargoNameEngTextBox.Text = ds.Tables["Cargo"].Select("Name = '" + cargoNameСomboBox.Text + "'")[0]["NameENG"].ToString();
            }
            if(cargoNameСomboBox.SelectedIndex == -1)
            {
                cargoNameEngTextBox.Text = "";
                cargoNameСomboBox.Text = "";
            }
        }

        private void firstPoint_TextChanged(object sender, EventArgs e)
        {
            if (firstPoint.Text != "" && vr != null)
            {
                secondPont.Text = (Convert.ToInt32(firstPoint.Text) + Convert.ToInt32(vr.Values[0][0]) - 1).ToString();
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

        private void addSender_Click(object sender, EventArgs e)
        {
            if (ds.Tables["Sender"].Select("Name = '" + senderNameRusTextBox.Text.ToUpper() + "'").Length == 0)
            {
                CheckFields();
                if (!invalidField)
                {
                    ChangeFieldsColor();
                    using (SqlCeConnection conn = new SqlCeConnection(sqlConnection))
                    {
                        conn.Open();
                        SqlCeCommand cmd = new SqlCeCommand("INSERT INTO Sender(ShortName, Name, Address, Cod, NameENG, AddressENG) VALUES (@ShortName, @Name, @Address, @Cod, @NameENG, @AddressENG)", conn);
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlCeParameter("ShortName", senderComboBox.Text));
                        cmd.Parameters.Add(new SqlCeParameter("Name", senderNameRusTextBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("Address", senderAddressRusTextBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("Cod", senderCodTextBox.Text));
                        cmd.Parameters.Add(new SqlCeParameter("NameENG", senderAddressEngTextBox.Text));
                        cmd.Parameters.Add(new SqlCeParameter("AddressENG", senderAddressEngTextBox.Text));
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
                            SqlCeCommand cmd = new SqlCeCommand("INSERT INTO Cargo(Name, Cod, ShortName, NameENG) VALUES (@Name, @Cod, @ShortName, @NameENG)", conn);
                            cmd.Parameters.Clear();
                            cmd.Parameters.Add(new SqlCeParameter("Name", cargoNameСomboBox.Text.ToUpper()));
                            cmd.Parameters.Add(new SqlCeParameter("Cod", cargoCodComboBox.Text));
                            cmd.Parameters.Add(new SqlCeParameter("ShortName", cargoShortName.Text.ToUpper()));
                            cmd.Parameters.Add(new SqlCeParameter("NameENG", cargoNameEngTextBox.Text.ToUpper()));
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
                        SqlCeCommand cmd = new SqlCeCommand("INSERT INTO PortT(Name, NameRus) VALUES (@Name, @NameRus)", conn);
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlCeParameter("Name", potComboBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("NameRus", potRusTextBox.Text.ToUpper()));
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
                        SqlCeCommand cmd = new SqlCeCommand("INSERT INTO CountryPort(Name, Port, PortRus) VALUES (@Name, @Port, @PortRus)", conn);
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlCeParameter("Name", countryСomboBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("Port", podComboBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("PortRus", podRusTextBox.Text.ToUpper()));
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
                        SqlCeCommand cmd = new SqlCeCommand("INSERT INTO Vessels(Name, Flag, FlagRus) VALUES (@Name, @Flag, @FlagRus)", conn);
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SqlCeParameter("Name", vesselsComboBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("Flag", flagEngTextBox.Text.ToUpper()));
                        cmd.Parameters.Add(new SqlCeParameter("FlagRus", flagRusTextBox.Text.ToUpper()));
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
            if (String.IsNullOrEmpty(voyageTextBox.Text))
            {
                voyageTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(flagEngTextBox.Text))
            {
                flagEngTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(flagRusTextBox.Text))
            {
                flagRusTextBox.BackColor = System.Drawing.Color.LightCoral;
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
            if (String.IsNullOrEmpty(potRusTextBox.Text))
            {
                potRusTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(podComboBox.Text))
            {
                podComboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(podRusTextBox.Text))
            {
                podRusTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(cargoNameСomboBox.Text))
            {
                cargoNameСomboBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(cargoNameEngTextBox.Text))
            {
                cargoNameEngTextBox.BackColor = System.Drawing.Color.LightCoral;
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
            if (String.IsNullOrEmpty(senderNameRusTextBox.Text))
            {
                senderNameRusTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(senderNameEngTextBox.Text))
            {
                senderNameEngTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(senderAddressRusTextBox.Text))
            {
                senderAddressRusTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(senderAddressEngTextBox.Text))
            {
                senderAddressEngTextBox.BackColor = System.Drawing.Color.LightCoral;
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
            if (String.IsNullOrEmpty(lineInfRusTextBox.Text))
            {
                lineInfRusTextBox.BackColor = System.Drawing.Color.LightCoral;
                invalidField = true;
            }
            if (String.IsNullOrEmpty(lineInfEngTextBox.Text))
            {
                lineInfEngTextBox.BackColor = System.Drawing.Color.LightCoral;
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
            voyageTextBox.BackColor = System.Drawing.Color.White;
            flagEngTextBox.BackColor = System.Drawing.Color.White;
            flagRusTextBox.BackColor = System.Drawing.Color.White;
            countryСomboBox.BackColor = System.Drawing.Color.White;
            potComboBox.BackColor = System.Drawing.Color.White;
            potRusTextBox.BackColor = System.Drawing.Color.White;
            podComboBox.BackColor = System.Drawing.Color.White;
            podRusTextBox.BackColor = System.Drawing.Color.White;
            cargoNameСomboBox.BackColor = System.Drawing.Color.White;
            cargoNameEngTextBox.BackColor = System.Drawing.Color.White;
            cargoCodComboBox.BackColor = System.Drawing.Color.White;
            contSizeComboBox.BackColor = System.Drawing.Color.White;
            contTypeComboBox.BackColor = System.Drawing.Color.White;
            senderNameRusTextBox.BackColor = System.Drawing.Color.White;
            senderNameEngTextBox.BackColor = System.Drawing.Color.White;
            senderAddressRusTextBox.BackColor = System.Drawing.Color.White;
            senderAddressEngTextBox.BackColor = System.Drawing.Color.White;
            senderCodTextBox.BackColor = System.Drawing.Color.White;
            receiverAddressTextBox.BackColor = System.Drawing.Color.White;
            lineInfRusTextBox.BackColor = System.Drawing.Color.White;
            lineInfEngTextBox.BackColor = System.Drawing.Color.White;
            firstPoint.BackColor = System.Drawing.Color.White;
            secondPont.BackColor = System.Drawing.Color.White;
        }

        private void ClearFields()
        {
            errendNumber.Text = "";
            linesComboBox.SelectedIndex = -1;
            linesComboBox.Text = "";
            senderComboBox.SelectedIndex = -1;           
            receiverComboBox.DataSource = new DataView(ds.Tables["Receiver"]).ToTable(true, "Name");
            receiverComboBox.SelectedIndex = -1;
            countryСomboBox.SelectedIndex = -1;
            countryСomboBox.Text = "";
            vesselsComboBox.SelectedIndex = -1;
            vesselsComboBox.Text = "";
            potComboBox.SelectedIndex = -1;
            potComboBox.Text = "";
            podComboBox.SelectedIndex = -1;
            podComboBox.Text = "";
            cargoNameСomboBox.SelectedIndex = -1;
            cargoNameСomboBox.Text = "";
            cargoNameEngTextBox.Text = "";
            cargoCodComboBox.SelectedIndex = cargoNameСomboBox.SelectedIndex;
            bookingTextBox.Text = "";
            contSizeComboBox.SelectedIndex = -1;
            contTypeComboBox.SelectedIndex = -1;
            firstPoint.Text = "";
            secondPont.Text = "";
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form f = new Form2();
            if (f.ShowDialog() == DialogResult.Cancel)
            {
                using (Process myProcess = new Process())
                {

                    System.Windows.Forms.Application.Exit();
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
                    System.Windows.Forms.Application.Exit();
                    Process.Start("Errend.exe", "");
                }
            }
        }

        private void clFields_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            ClearFields();
        }

        private void cTOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
            Thread tr = new Thread(OpenForm);
            tr.SetApartmentState(ApartmentState.STA);
            tr.Start();
        }
        
        private void OpenForm()
        {
            System.Windows.Forms.Application.Run(new Form1());
        }

        private void CreateXLS(ValueRange vr, params string[] data)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            Workbook worKbooK = excel.Workbooks.Add(Type.Missing);
            Worksheet worksheet = worKbooK.ActiveSheet;
            worksheet.Name = "Table";
            Range celLrangE;
            worksheet.Range["B1:B9"].NumberFormat = "@";
            for (int i = 0; i < 9; i++)
            {
                worksheet.Range[worksheet.Cells[i + 1, 2], worksheet.Cells[i + 1, 8]].Merge();
                worksheet.Cells[i + 1, 1] = tableCMA[i].ToUpper();
                worksheet.Cells[i + 1, 2] = data[i].ToUpper();
                worksheet.Cells[i + 1, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Cells[i + 1, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Cells[i + 1, 2].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Cells[i + 1, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;              
            }

            worksheet.Cells[10, 1].Formula = tableCMA[9];
            worksheet.Cells[10, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.Cells[10, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;

            for (int i = 2; i < 9; i++)
            {
                worksheet.Cells[10, i] = tableCMA[i + 8];
                worksheet.Cells[10, i].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                worksheet.Cells[10, i].VerticalAlignment = XlVAlign.xlVAlignCenter;
            }

            for (int i = 0; i < vr.Values.Count; i++)
            {
                worksheet.Cells[i + 11, 1] = i + 1;
                worksheet.Cells[i + 11, 3] = vr.Values[i][0];
                worksheet.Cells[i + 11, 4] = contSizeComboBox.Text;
                worksheet.Cells[i + 11, 5] = contTypeComboBox.Text;
                worksheet.Cells[i + 11, 6] = vr.Values[i][13];
                worksheet.Cells[i + 11, 7] = vr.Values[i][4];
                worksheet.Cells[i + 11, 8] = vr.Values[i][19];
                worksheet.Range["A11:H100"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                worksheet.Range["A11:H100"].VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
            worksheet.Range["A1:A9"].Interior.Color = System.Drawing.Color.Gray;
            worksheet.Range["B1:B9"].Interior.Color = System.Drawing.Color.LightGray;
            worksheet.Range["A10"].Interior.Color = System.Drawing.Color.Yellow;
            worksheet.Range["B10:H10"].Interior.Color = System.Drawing.Color.LightBlue;
            celLrangE = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[10 + vr.Values.Count, 8]];
            celLrangE.EntireColumn.AutoFit();
            worksheet.Range[worksheet.Cells[6, 2], worksheet.Cells[7, 2]].EntireRow.RowHeight = 80;
            Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            worKbooK.SaveAs(@"" + Properties.Settings.Default["SavingPath"].ToString() + "\\" + linesComboBox.Text + "\\" + vesselsComboBox.Text + "\\" + senderComboBox.Text + " - Таб. " + errendNumber.Text + " - " + vr.Values.Count.ToString() + " конт. - " + countryСomboBox.Text + ".xlsx");
            worKbooK.Close();
            excel.Quit();
            Process.Start(@"" + Properties.Settings.Default["SavingPath"].ToString() + "\\" + linesComboBox.Text + "\\" + vesselsComboBox.Text + "\\" + senderComboBox.Text + " - Таб. " + errendNumber.Text + " - " + vr.Values.Count.ToString() + " конт. - " + countryСomboBox.Text + ".xlsx");
        }

        private void firstPoint_TextChanged_1(object sender, EventArgs e)
        {
            if (firstPoint.Text != "" && vr != null)
            {
                secondPont.Text = (Convert.ToInt32(firstPoint.Text) + Convert.ToInt32(vr.Values[0][0]) - 1).ToString();
            }
        }

        private void potComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (potComboBox.SelectedIndex != -1)
            {
                potRusTextBox.Text = ds.Tables["PortT"].Select("Name = '" + potComboBox.Text + "'")[0]["NameRus"].ToString();
            }
            else
            {
                potRusTextBox.Text = "";
            }
        }

        private void podComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (podComboBox.SelectedIndex != -1)
            {
                podRusTextBox.Text = ds.Tables["CountryPort"].Select("Port = '" + podComboBox.Text + "'")[0]["PortRus"].ToString();
            }
            else
            {
                podRusTextBox.Text = "";
            }
        }

        private void Arena_CheckedChanged(object sender, EventArgs e)
        {
            if (arena.Checked == true)
            {
                ugl.Checked = false;
            }
        }

        private void Ugl_CheckedChanged(object sender, EventArgs e)
        {
            if (ugl.Checked == true)
            {
                arena.Checked = false;
            }
        }
    }
}
