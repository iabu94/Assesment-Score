using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;

namespace AssesmentScore
{
    public partial class Main : Form
    {
        private double TotalEconomic = 0;
        private double TotalEnvironmental = 0;
        private double TotalSocial = 0;
        private double Total = 0;
        private int NextValue = 0;

        public Main()
        {
            InitializeComponent();

            //this.WindowState = FormWindowState.Maximized;

            panelMoving.Visible = false;
            checkBox11.Visible = false;
            checkBox21.Visible = false;
            checkBox31.Visible = false;
            checkBox12.Visible = false;
            checkBox22.Visible = false;
            checkBox32.Visible = false;
            checkBox13.Visible = false;
            checkBox23.Visible = false;
            checkBox33.Visible = false;
            checkBox14.Visible = false;
            checkBox24.Visible = false;
            checkBox34.Visible = false;
            checkBox15.Visible = false;
            checkBox25.Visible = false;
            checkBox16.Visible = false;

            panelRight.Visible = false;
            panelBottom.Visible = false;

        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == WM_NCHITTEST)
                m.Result = (IntPtr)(HT_CAPTION);
        }

        private const int WM_NCHITTEST = 0x84;
        private const int HT_CLIENT = 0x1;
        private const int HT_CAPTION = 0x2;

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        #region Side button clicks

        private void buttonEconomic_Click(object sender, EventArgs e)
        {
            NextValue = 1;
            panelBottom.Visible = true;
            panelRight.Visible = true;
            panelMoving.Visible = true;
            panelMoving.Height = buttonEconomic.Height;
            panelMoving.Top = buttonEconomic.Top;

            VisibleGroup1();
            NonVisibleGroup2();
            NonVisibleGroup3();

            economicUserControl1.BringToFront();

            labelPoints.Text = CalcTotalEconomic().ToString();
        }

        private void buttonEnvironment_Click(object sender, EventArgs e)
        {
            NextValue = 2;
            panelBottom.Visible = true;

            panelRight.Visible = true;
            panelMoving.Visible = true;
            panelMoving.Height = buttonEnvironment.Height;
            panelMoving.Top = buttonEnvironment.Top;

            VisibleGroup2();
            NonVisibleGroup1();
            NonVisibleGroup3();

            environmentalUserControl1.BringToFront();

            labelPoints.Text = CalcTotalEnvironment().ToString();
        }

        private void buttonSocial_Click(object sender, EventArgs e)
        {
            NextValue = 3;
            panelBottom.Visible = true;

            panelRight.Visible = true;
            panelMoving.Visible = true;
            panelMoving.Height = buttonSocial.Height;
            panelMoving.Top = buttonSocial.Top;

            VisibleGroup3();
            NonVisibleGroup1();
            NonVisibleGroup2();

            socialUserControl1.BringToFront();

            // TotalSocial = 0;
            labelPoints.Text = CalcTotalSocial().ToString();
        }

        #endregion

        #region Visible and nin visible groups

        private void NonVisibleGroup1()
        {
            checkBox11.Visible = false;
            checkBox12.Visible = false;
            checkBox13.Visible = false;
            checkBox14.Visible = false;
            checkBox15.Visible = false;
            checkBox16.Visible = false;
        }

        private void NonVisibleGroup2()
        {
            checkBox21.Visible = false;
            checkBox22.Visible = false;
            checkBox23.Visible = false;
            checkBox24.Visible = false;
            checkBox25.Visible = false;
        }

        private void NonVisibleGroup3()
        {
            checkBox31.Visible = false;
            checkBox32.Visible = false;
            checkBox33.Visible = false;
            checkBox34.Visible = false;
        }

        private void VisibleGroup1()
        {
            checkBox11.Visible = true;
            checkBox12.Visible = true;
            checkBox13.Visible = true;
            checkBox14.Visible = true;
            checkBox15.Visible = true;
            checkBox16.Visible = true;
        }

        private void VisibleGroup2()
        {
            checkBox21.Visible = true;
            checkBox22.Visible = true;
            checkBox23.Visible = true;
            checkBox24.Visible = true;
            checkBox25.Visible = true;
        }

        private void VisibleGroup3()
        {
            checkBox31.Visible = true;
            checkBox32.Visible = true;
            checkBox33.Visible = true;
            checkBox34.Visible = true;
        }

        #endregion

        #region Checkbox check events

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalEconomic().ToString();
        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalEconomic().ToString();
        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalEconomic().ToString();
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalEconomic().ToString();
        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalEconomic().ToString();
        }

        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalEconomic().ToString();
        }

        private void checkBox21_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalEnvironment().ToString();
        }

        private void checkBox22_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalEnvironment().ToString();
        }

        private void checkBox23_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalEnvironment().ToString();
        }

        private void checkBox24_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalEnvironment().ToString();
        }

        private void checkBox25_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalEnvironment().ToString();
        }

        private void checkBox31_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalSocial().ToString();
        }

        private void checkBox32_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalSocial().ToString();
        }

        private void checkBox33_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalSocial().ToString();
        }

        private void checkBox34_CheckedChanged(object sender, EventArgs e)
        {
            labelPoints.Text = CalcTotalSocial().ToString();
        }

        #endregion

        #region Calcualte Points

        private double CalcTotalEconomic()
        {
            TotalEconomic = 0;
            if (checkBox11.Checked)
            {
                TotalEconomic += 42.89;
            }

            if (checkBox12.Checked)
            {
                TotalEconomic += 17.44;
            }
            if (checkBox13.Checked)
            {
                TotalEconomic += 17.00;
            }
            if (checkBox14.Checked)
            {
                TotalEconomic += 12.02;
            }
            if (checkBox15.Checked)
            {
                TotalEconomic += 6.23;
            }
            if (checkBox16.Checked)
            {
                TotalEconomic += 4.42;
            }
            return TotalEconomic;
        }

        private double CalcTotalEnvironment()
        {
            TotalEnvironmental = 0;
            if (checkBox21.Checked)
            {
                TotalEnvironmental += 41.63;
            }

            if (checkBox22.Checked)
            {
                TotalEnvironmental += 26.93;
            }
            if (checkBox23.Checked)
            {
                TotalEnvironmental += 12.83;
            }
            if (checkBox24.Checked)
            {
                TotalEnvironmental += 12.37;
            }
            if (checkBox25.Checked)
            {
                TotalEnvironmental += 6.22;
            }
            return TotalEnvironmental;
        }

        private double CalcTotalSocial()
        {
            TotalSocial = 0;
            if (checkBox31.Checked)
            {
                TotalSocial += 35.69;
            }

            if (checkBox32.Checked)
            {
                TotalSocial += 23.38;
            }
            if (checkBox33.Checked)
            {
                TotalSocial += 23.05;
            }
            if (checkBox34.Checked)
            {
                TotalSocial += 17.88;
            }
            return TotalSocial;
        }

        #endregion

        private void button5_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void buttonHome_Click(object sender, EventArgs e)
        {
            panelBottom.Visible = false;
            panelRight.Visible = false;
            pictureBoxUserControl1.BringToFront();
            //pictureBoxHome.Visible = true;
            //panelPictureBox.Visible = true;
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dtblEcon = MakeDataTableEconomic();
                DataTable dtblEnv = MakeDataTableEnvironmental();
                DataTable dtblSoc = MakeDataTableSocial();
                var folderBrowserDialog1 = new FolderBrowserDialog();

                // Show the FolderBrowserDialog.
                DialogResult result = folderBrowserDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    string folderName = folderBrowserDialog1.SelectedPath;
                    ExportDataTableToPdf(dtblEcon, dtblEnv, dtblSoc, folderName + "Assesment Score.pdf", "SUSTAINABILTY ASSESSMENT OF PREFABRICATION: BUILDING CONSTRUCTION PROJECT IN SRI LANKA");
                    MessageBox.Show("Save Successfully as Assesment Score.pdf");

                    //open pdf
                    DialogResult res = MessageBox.Show("Do you want to open the exported pdf..?", "Open File", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (res == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(folderName + "Assesment Score.pdf");
                        this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Message");
            }
        }

        #region MakeDataTables
        DataTable MakeDataTableEconomic()
        {
            //Create friend table object
            DataTable pointsOwned = new DataTable();

            //Define columns
            pointsOwned.Columns.Add("Criteria Description");
            pointsOwned.Columns.Add("Points Allocated");
            pointsOwned.Columns.Add("Yes/No");
            //pointsOwned.Columns.Add("Region");

            //Assigning the check box states
            #region YES/NO
            string checkCheckBox11 = "NO";
            string checkCheckBox12 = "NO";
            string checkCheckBox13 = "NO";
            string checkCheckBox14 = "NO";
            string checkCheckBox15 = "NO";
            string checkCheckBox16 = "NO";
            if (checkBox11.Checked)
            {
                checkCheckBox11 = "YES";
            }
            if (checkBox12.Checked)
            {
                checkCheckBox12 = "YES";
            }
            if (checkBox13.Checked)
            {
                checkCheckBox13 = "YES";
            }
            if (checkBox14.Checked)
            {
                checkCheckBox14 = "YES";
            }
            if (checkBox15.Checked)
            {
                checkCheckBox15 = "YES";
            }
            if (checkBox16.Checked)
            {
                checkCheckBox16 = "YES";
            }
            #endregion

            //Populate with points :)
            pointsOwned.Rows.Add("Cost efficient ", "42.89", checkCheckBox11);
            pointsOwned.Rows.Add("Quality improvement ", "17.44", checkCheckBox12);
            pointsOwned.Rows.Add("Increasing Construction efficiency", "17.00", checkCheckBox13);
            pointsOwned.Rows.Add("Resource efficiency", "12.02", checkCheckBox14);
            pointsOwned.Rows.Add("Less wastage", "6.23", checkCheckBox15);
            pointsOwned.Rows.Add("Time saving", "4.42", checkCheckBox16);
            pointsOwned.Rows.Add("Total", "100", CalcTotalEconomic().ToString());

            return pointsOwned;
        }
        DataTable MakeDataTableEnvironmental()
        {
            //Create friend table object
            DataTable pointsOwned = new DataTable();

            //Define columns
            pointsOwned.Columns.Add("Criteria Description");
            pointsOwned.Columns.Add("Points Allocated");
            pointsOwned.Columns.Add("Yes/No");

            //Assign the YES/NO
            #region Assign checkbox YES/NO

            string checkCheckBox21 = "NO";
            string checkCheckBox22 = "NO";
            string checkCheckBox23 = "NO";
            string checkCheckBox24 = "NO";
            string checkCheckBox25 = "NO";
            if (checkBox21.Checked)
            {
                checkCheckBox21 = "YES";
            }
            if (checkBox22.Checked)
            {
                checkCheckBox22 = "YES";
            }
            if (checkBox23.Checked)
            {
                checkCheckBox23 = "YES";
            }
            if (checkBox24.Checked)
            {
                checkCheckBox24 = "YES";
            }
            if (checkBox25.Checked)
            {
                checkCheckBox25 = "YES";
            }

            #endregion

            //Populate with points :)
            pointsOwned.Rows.Add("Reduction of environmental pollution", "41.63", checkCheckBox21);
            pointsOwned.Rows.Add("Waste management practices", "26.93", checkCheckBox22);
            pointsOwned.Rows.Add("Energy saving", "12.83", checkCheckBox23);
            pointsOwned.Rows.Add("Ability to reuse/ recyclable material", "12.37", checkCheckBox24);
            pointsOwned.Rows.Add("Water conservation", "6.22", checkCheckBox25);
            pointsOwned.Rows.Add("Total", "100", CalcTotalEnvironment().ToString());

            return pointsOwned;
        }
        DataTable MakeDataTableSocial()
        {
            //Create friend table object
            DataTable pointsOwned = new DataTable();

            //Define columns
            pointsOwned.Columns.Add("Criteria Description");
            pointsOwned.Columns.Add("Points Allocated");
            pointsOwned.Columns.Add("Yes/No");
            //pointsOwned.Columns.Add("Region");

            string checkCheckBox31 = "NO";
            string checkCheckBox32 = "NO";
            string checkCheckBox33 = "NO";
            string checkCheckBox34 = "NO";
            if (checkBox31.Checked)
            {
                checkCheckBox31 = "YES";
            }
            if (checkBox32.Checked)
            {
                checkCheckBox32 = "YES";
            }
            if (checkBox33.Checked)
            {
                checkCheckBox33 = "YES";
            }
            if (checkBox34.Checked)
            {
                checkCheckBox34 = "YES";
            }

            //Populate with points :)
            pointsOwned.Rows.Add("Increased workers efficiency", "35.69", checkCheckBox31);
            pointsOwned.Rows.Add("Increased certainty and less risk", "23.38", checkCheckBox32);
            pointsOwned.Rows.Add("Health and Safety improvement", "23.05", checkCheckBox33);
            pointsOwned.Rows.Add("Less site congestion", "17.88", checkCheckBox34);
            pointsOwned.Rows.Add("Total", "100", CalcTotalSocial().ToString());
            //pointsOwned.Rows.Add("6", "Khalid", "UAE", "Dubai");
            //pointsOwned.Rows.Add("7", "William", "Australia", "Canberra");

            return pointsOwned;
        }

        #endregion

        private void ExportDataTableToPdf(DataTable dtblEconomic, DataTable dtblEnvironment, DataTable dtblSocial, String strPdfPath, string strHeader)
        {
            System.IO.FileStream fs = new FileStream(strPdfPath, FileMode.Create, FileAccess.Write, FileShare.None);
            Document document = new Document();
            document.SetPageSize(iTextSharp.text.PageSize.A4);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();

            //Report Header
            BaseFont bfntHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntHead = new Font(bfntHead, 16, 1, Color.BLUE);
            Paragraph prgHeading = new Paragraph();
            prgHeading.Alignment = Element.ALIGN_CENTER;
            prgHeading.Add(new Chunk(strHeader.ToUpper(), fntHead));
            document.Add(prgHeading);

            //Author
            Paragraph prgAuthor = new Paragraph();
            BaseFont btnAuthor = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntAuthor = new Font(btnAuthor, 8, 2, Color.GRAY);
            prgAuthor.Alignment = Element.ALIGN_RIGHT;
            //prgAuthor.Add(new Chunk("Author : Fathima Aqsha", fntAuthor));
            prgAuthor.Add(new Chunk("\nDate : " + DateTime.Now.ToShortDateString(), fntAuthor));
            document.Add(prgAuthor);

            //Add a line seperation
            Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, Color.BLACK, Element.ALIGN_LEFT, 1)));
            document.Add(p);

            #region Economic Table PDF
            //Table Heading
            BaseFont bfntSubHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntSubHead = new Font(bfntSubHead, 12, 3, Color.BLACK);
            //Add line break
            document.Add(new Chunk("\n", fntSubHead));
            Paragraph prgSubHeading = new Paragraph();
            prgSubHeading.Alignment = Element.ALIGN_CENTER;
            prgSubHeading.Add(new Chunk("Economic Indicator - 45.32%", fntSubHead));
            document.Add(prgSubHeading);

            //Add line break
            Font fontLineBreak = new Font(bfntSubHead, 5, 1, Color.GRAY);
            document.Add(new Chunk("\n", fontLineBreak));

            //Write the table
            PdfPTable tableEC = new PdfPTable(dtblEconomic.Columns.Count);
            //Table header
            BaseFont btnColumnHeader = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntColumnHeader = new Font(btnColumnHeader, 10, 1, Color.WHITE);
            for (int i = 0; i < dtblEconomic.Columns.Count; i++)
            {
                PdfPCell cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                cell.AddElement(new Chunk(dtblEconomic.Columns[i].ColumnName.ToUpper(), fntColumnHeader));
                tableEC.AddCell(cell);
            }
            //table Data
            for (int i = 0; i < dtblEconomic.Rows.Count; i++)
            {
                for (int j = 0; j < dtblEconomic.Columns.Count; j++)
                {
                    tableEC.AddCell(dtblEconomic.Rows[i][j].ToString());
                }
            }

            document.Add(tableEC);
            #endregion

            #region Environment Table PDF
            //Add a line break
            document.Add(new Chunk("\n", fontLineBreak));
            //Table Heading
            Paragraph prgSubHeading1 = new Paragraph();
            prgSubHeading1.Alignment = Element.ALIGN_CENTER;
            prgSubHeading1.Add(new Chunk("Environmental Indicator - 41.58%", fntSubHead));
            document.Add(prgSubHeading1);

            //Add line break
            document.Add(new Chunk("\n", fontLineBreak));

            //Write the table
            PdfPTable tableEN = new PdfPTable(dtblEnvironment.Columns.Count);
            //Table header
            for (int i = 0; i < dtblEnvironment.Columns.Count; i++)
            {
                PdfPCell cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                cell.AddElement(new Chunk(dtblEnvironment.Columns[i].ColumnName.ToUpper(), fntColumnHeader));
                tableEN.AddCell(cell);
            }
            //table Data
            for (int i = 0; i < dtblEnvironment.Rows.Count; i++)
            {
                for (int j = 0; j < dtblEnvironment.Columns.Count; j++)
                {
                    tableEN.AddCell(dtblEnvironment.Rows[i][j].ToString());
                }
            }

            document.Add(tableEN);
            #endregion

            #region Social Table PDF

            //Table Heading
            //Add line break
            document.Add(new Chunk("\n", fntSubHead));
            Paragraph prgSubHeading2 = new Paragraph();
            prgSubHeading2.Alignment = Element.ALIGN_CENTER;
            prgSubHeading2.Add(new Chunk("Social Indicator - 13.10%", fntSubHead));
            document.Add(prgSubHeading2);

            //Add line break
            document.Add(new Chunk("\n", fontLineBreak));

            //Write the table
            PdfPTable tableSC = new PdfPTable(dtblSocial.Columns.Count);
            //Table header
            for (int i = 0; i < dtblSocial.Columns.Count; i++)
            {
                PdfPCell cell = new PdfPCell();
                cell.BackgroundColor = Color.GRAY;
                cell.AddElement(new Chunk(dtblSocial.Columns[i].ColumnName.ToUpper(), fntColumnHeader));
                tableSC.AddCell(cell);
            }
            //table Data
            for (int i = 0; i < dtblSocial.Rows.Count; i++)
            {
                for (int j = 0; j < dtblSocial.Columns.Count; j++)
                {
                    tableSC.AddCell(dtblSocial.Rows[i][j].ToString());
                }
            }

            document.Add(tableSC);

            #endregion

            document.Add(new Chunk("\n", fntSubHead));

            Total = 0;
            Total = CalcTotalEconomic() * 45.32;
            Total += CalcTotalEnvironment() * 41.58;
            Total += CalcTotalSocial() * 13.10;
            Total = Total / 100;


            BaseFont btnFinalResult = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntFinalResult = new Font(btnFinalResult, 12, 1, Color.BLACK);
            Paragraph finalResult = new Paragraph();
            finalResult.Alignment = Element.ALIGN_CENTER;
            finalResult.Add(new Chunk("Sustainability achievement of the prefabrication project - " + Math.Round(Total, 2) + " %", fntFinalResult));
            document.Add(finalResult);


            document.Close();
            writer.Close();
            fs.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void buttonNext_Click(object sender, EventArgs e)
        {
            if (NextValue == 1)
            {
                buttonEnvironment.PerformClick();
            }
            else if (NextValue == 2)
            {
                buttonSocial.PerformClick();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (NextValue == 2)
            {
                buttonEconomic.PerformClick();
            }
            else if (NextValue == 3)
            {
                buttonEnvironment.PerformClick();
            }
        }
    }
}
