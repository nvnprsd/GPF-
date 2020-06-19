using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Data.OleDb;



namespace GPF
{
    public partial class GPFcal : Form
    {
        public GPFcal()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
           
        }
        void cal()
        {
            APR_FD.Text = (Convert.ToDouble(Apr.Text) + Convert.ToDouble(apr_b.Text) + Convert.ToDouble(apr_r.Text)).ToString();
            MAY_FD.Text = (Convert.ToDouble(May.Text) + Convert.ToDouble(may_b.Text) + Convert.ToDouble(may_r.Text)).ToString();
            JUN_FD.Text = (Convert.ToDouble(Jun.Text) + Convert.ToDouble(jun_b.Text) + Convert.ToDouble(jun_r.Text)).ToString();
            JUL_FD.Text = (Convert.ToDouble(Jul.Text) + Convert.ToDouble(jul_b.Text) + Convert.ToDouble(jul_r.Text)).ToString();
            AUG_FD.Text = (Convert.ToDouble(Aug.Text) + Convert.ToDouble(aug_b.Text) + Convert.ToDouble(aug_r.Text)).ToString();
            SEP_FD.Text = (Convert.ToDouble(Sep.Text) + Convert.ToDouble(sep_b.Text) + Convert.ToDouble(sep_r.Text)).ToString();
            OCT_FD.Text = (Convert.ToDouble(Oct.Text) + Convert.ToDouble(oct_b.Text) + Convert.ToDouble(oct_r.Text)).ToString();
            NOV_FD.Text = (Convert.ToDouble(Nov.Text) + Convert.ToDouble(nov_b.Text) + Convert.ToDouble(nov_r.Text)).ToString();
            DEC_FD.Text = (Convert.ToDouble(Dec.Text) + Convert.ToDouble(dec_b.Text) + Convert.ToDouble(dec_r.Text)).ToString();
            JAN_FD.Text = (Convert.ToDouble(Jan.Text) + Convert.ToDouble(jan_b.Text) + Convert.ToDouble(jan_r.Text)).ToString();
            FEB_FD.Text = (Convert.ToDouble(Feb.Text) + Convert.ToDouble(feb_b.Text) + Convert.ToDouble(feb_r.Text)).ToString();
            MAR_FD.Text = (Convert.ToDouble(Mar.Text) + Convert.ToDouble(mar_b.Text) + Convert.ToDouble(mar_r.Text)).ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try {
              
                if (Name.Text == "" ||school.Text=="")
                {
                    MessageBox.Show("Name or School Name Must be Filled");
                }
                else
                {
                    if (radioButton1.Checked == false && radioButton2.Checked == false && radioButton3.Checked == false)
                    {
                        MessageBox.Show("Interest Rate Must Be selected Annually, Quaterly or Monthly", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    
                    else
                    {
                        double totalintr = 0;
                        MessageBox.Show("Let's Get The Results...!");
                        double a = (Convert.ToDouble(opening_bal.Text) + Convert.ToDouble(APR_FD.Text)) - Convert.ToDouble(final_wd_apr.Text);
                        close_apr.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m1.Text))) / 1200);

                        a = Convert.ToDouble(close_apr.Text) + Convert.ToDouble(MAY_FD.Text) - Convert.ToDouble(final_wd_may.Text);
                        close_may.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m2.Text))) / 1200);
                        a = Convert.ToDouble(close_may.Text) + Convert.ToDouble(JUN_FD.Text) - Convert.ToDouble(final_wd_jun.Text);
                        close_jun.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m3.Text))) / 1200);
                        a = Convert.ToDouble(close_jun.Text) + Convert.ToDouble(JUL_FD.Text) - Convert.ToDouble(final_wd_jul.Text);
                        close_jul.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m4.Text))) / 1200);
                        a = Convert.ToDouble(close_jul.Text) + Convert.ToDouble(AUG_FD.Text) - Convert.ToDouble(final_wd_aug.Text);
                        close_aug.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m5.Text))) / 1200);
                        a = Convert.ToDouble(close_aug.Text) + Convert.ToDouble(SEP_FD.Text) - Convert.ToDouble(final_wd_sep.Text);
                        close_sep.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m6.Text))) / 1200);
                        a = Convert.ToDouble(close_sep.Text) + Convert.ToDouble(OCT_FD.Text) - Convert.ToDouble(final_wd_oct.Text);
                        close_oct.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m7.Text))) / 1200);

                        a = Convert.ToDouble(close_oct.Text) + Convert.ToDouble(NOV_FD.Text) - Convert.ToDouble(final_wd_nov.Text);
                        close_nov.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m8.Text))) / 1200);

                        a = Convert.ToDouble(close_nov.Text) + Convert.ToDouble(DEC_FD.Text) - Convert.ToDouble(final_wd_dec.Text);
                        close_dec.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m9.Text))) / 1200);

                        a = Convert.ToDouble(close_dec.Text) + Convert.ToDouble(JAN_FD.Text) - Convert.ToDouble(final_wd_jan.Text);
                        close_jan.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m10.Text))) / 1200);

                        a = Convert.ToDouble(close_jan.Text) + Convert.ToDouble(FEB_FD.Text) - Convert.ToDouble(final_wd_feb.Text);
                        close_feb.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m11.Text))) / 1200);


                        a = Convert.ToDouble(close_feb.Text) + Convert.ToDouble(MAR_FD.Text) - Convert.ToDouble(final_wd_mar.Text);
                        close_mar.Text = a.ToString();
                        totalintr += ((a * (Convert.ToDouble(m12.Text))) / 1200);
                        intt.Text = Math.Round(totalintr, 2).ToString();
                        closing_bal.Text = Math.Round((totalintr + (Convert.ToDouble(close_mar.Text))), 2).ToString();

                        sum_subcr.Text = (Convert.ToDouble(Apr.Text) +   Convert.ToDouble(May.Text) + Convert.ToDouble(Jun.Text) + Convert.ToDouble(Jul.Text) + Convert.ToDouble(Aug.Text) + Convert.ToDouble(Sep.Text) + Convert.ToDouble(Oct.Text) + Convert.ToDouble(Nov.Text) + Convert.ToDouble(Dec.Text) + Convert.ToDouble(Jan.Text) + Convert.ToDouble(Feb.Text) + Convert.ToDouble(Mar.Text)).ToString();
                        sum_bonus.Text = (Convert.ToDouble(apr_b.Text) + Convert.ToDouble(may_b.Text) + Convert.ToDouble(jun_b.Text) + Convert.ToDouble(jul_b.Text) + Convert.ToDouble(aug_b.Text) + Convert.ToDouble(sep_b.Text) + Convert.ToDouble(oct_b.Text) + Convert.ToDouble(nov_b.Text) + Convert.ToDouble(dec_b.Text) + Convert.ToDouble(jan_b.Text) + Convert.ToDouble(feb_b.Text) + Convert.ToDouble(mar_b.Text)).ToString();
                        sum_refund.Text = (Convert.ToDouble(apr_r.Text) + Convert.ToDouble(may_r.Text) + Convert.ToDouble(jun_r.Text) + Convert.ToDouble(jul_r.Text) + Convert.ToDouble(aug_r.Text) + Convert.ToDouble(sep_r.Text) + Convert.ToDouble(oct_r.Text) + Convert.ToDouble(nov_r.Text) + Convert.ToDouble(dec_r.Text) + Convert.ToDouble(jan_r.Text) + Convert.ToDouble(feb_r.Text) + Convert.ToDouble(mar_r.Text)).ToString();
                        Sum_Fdep.Text = (Convert.ToDouble(APR_FD.Text) + Convert.ToDouble(MAY_FD.Text) + Convert.ToDouble(JUN_FD.Text) + Convert.ToDouble(JUL_FD.Text) + Convert.ToDouble(AUG_FD.Text) + Convert.ToDouble(SEP_FD.Text) + Convert.ToDouble(OCT_FD.Text) + Convert.ToDouble(NOV_FD.Text) + Convert.ToDouble(DEC_FD.Text) + Convert.ToDouble(JAN_FD.Text) + Convert.ToDouble(FEB_FD.Text) + Convert.ToDouble(MAR_FD.Text)).ToString();
                        sum_tempad.Text = (Convert.ToDouble(apr_cr.Text) + Convert.ToDouble(may_cr.Text) + Convert.ToDouble(jun_cr.Text) + Convert.ToDouble(jul_cr.Text) + Convert.ToDouble(aug_cr.Text) + Convert.ToDouble(sep_cr.Text) + Convert.ToDouble(oct_cr.Text) + Convert.ToDouble(nov_cr.Text) + Convert.ToDouble(dec_cr.Text) + Convert.ToDouble(jan_cr.Text) + Convert.ToDouble(feb_cr.Text) + Convert.ToDouble(mar_cr.Text)).ToString();
                        sum_owith.Text = (Convert.ToDouble(with_apr.Text) + Convert.ToDouble(with_may.Text) + Convert.ToDouble(with_jun.Text) + Convert.ToDouble(with_jul.Text) + Convert.ToDouble(with_aug.Text) + Convert.ToDouble(with_sep.Text) + Convert.ToDouble(with_oct.Text) + Convert.ToDouble(with_nov.Text) + Convert.ToDouble(with_dec.Text) + Convert.ToDouble(with_jan.Text) + Convert.ToDouble(with_feb.Text) + Convert.ToDouble(with_mar.Text)).ToString();
                        sum_fwith.Text = (Convert.ToDouble(final_wd_apr.Text) + Convert.ToDouble(final_wd_may.Text) + Convert.ToDouble(final_wd_jun.Text) + Convert.ToDouble(final_wd_jul.Text) + Convert.ToDouble(final_wd_aug.Text) + Convert.ToDouble(final_wd_sep.Text) + Convert.ToDouble(final_wd_oct.Text) + Convert.ToDouble(final_wd_nov.Text) + Convert.ToDouble(final_wd_dec.Text) + Convert.ToDouble(final_wd_jan.Text) + Convert.ToDouble(final_wd_feb.Text) + Convert.ToDouble(final_wd_mar.Text)).ToString();
                        sum_closing.Text = (Convert.ToDouble(close_apr.Text) + Convert.ToDouble(close_may.Text) + Convert.ToDouble(close_jun.Text) + Convert.ToDouble(close_jul.Text) + Convert.ToDouble(close_aug.Text) + Convert.ToDouble(close_sep.Text) + Convert.ToDouble(close_oct.Text) + Convert.ToDouble(close_nov.Text) + Convert.ToDouble(close_dec.Text) + Convert.ToDouble(close_jan.Text) + Convert.ToDouble(close_feb.Text) + Convert.ToDouble(close_mar.Text)).ToString();

                        label4.Text = Math.Round(Convert.ToDouble(closing_bal.Text), 0).ToString();
                        label5.Text = "Rupees " + numtoword(Convert.ToInt32(label4.Text)).ToString() + " Only.";
                        label4.Text = label4.Text + "/-";
                        rem_bal.Text = (Convert.ToDouble(rem_bal.Text) + (Convert.ToDouble(sum_tempad.Text))).ToString();
                        rem_bal.Text = (Convert.ToDouble(rem_bal.Text) - Convert.ToDouble(sum_refund.Text)).ToString();

                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
         }
       
        private string numtoword(int number)
        {

            string words = "";
            if ((number / 1000000) > 0)
            {
                words += numtoword(number / 100000) + " LAKH";
                number %= 1000000;
            }
            if ((number / 1000) > 0)
            {
                words += numtoword(number / 1000) + " THOUSAND ";
                number %= 1000;
            }
            if ((number / 100) > 0)
            {
                words += numtoword(number / 100) + " HUNDRED ";
                number %= 100;
            }
            if (number > 0)
            {
                if (words != "") words += "AND ";
                var unitsMap = new[]
                {
            "ZERO", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE", "TEN", "ELEVEN", "TWELVE", "THIRTEEN", "FOURTEEN", "FIFTEEN", "SIXTEEN", "SEVENTEEN", "EIGHTEEN", "NINETEEN"
        };
                var tensMap = new[]
                {
            "ZERO", "TEN", "TWENTY", "THIRTY", "FORTY", "FIFTY", "SIXTY", "SEVENTY", "EIGHTY", "NINETY"
        };
                if (number < 20) words += unitsMap[number];
                else
                {
                    words += tensMap[number / 10];
                    if ((number % 10) > 0) words += " " + unitsMap[number % 10];
                }
            }
            return words;

        }

        private void label27_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        bool g = false;
        private void Apr_Leave(object sender, EventArgs e)
        {
            DialogResult r = MessageBox.Show("Want To add Auto Values","Alert",MessageBoxButtons.YesNo);


            if (r == DialogResult.Yes)
            {
                May.Text = Jun.Text = Jul.Text = Aug.Text = Sep.Text = Oct.Text = Nov.Text = Dec.Text = Jan.Text = Feb.Text = Mar.Text = Apr.Text;
                cal();
            }
            else {
                cal();
            }
        }
        private void mar_r_Leave2(object sender, EventArgs e)
        {
            try
            {
                DialogResult r = MessageBox.Show("Want To add Auto Values", "Alert", MessageBoxButtons.YesNo);


                if (r == DialogResult.Yes)
                {
                    may_r.Text = jun_r.Text = jul_r.Text = aug_r.Text = sep_r.Text = oct_r.Text = nov_r.Text = dec_r.Text = jan_r.Text = feb_r.Text = mar_r.Text = apr_r.Text;
                }
                cal();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        private void mar_r_Leave(object sender, EventArgs e)
        {
            try {
              
                cal();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void with_mar_Leave(object sender, EventArgs e)
        {
            try
            { double a = (Convert.ToDouble(apr_cr.Text) + Convert.ToDouble(may_cr.Text) + Convert.ToDouble(jun_cr.Text) + Convert.ToDouble(jul_cr.Text) + Convert.ToDouble(aug_cr.Text) + Convert.ToDouble(sep_cr.Text) + Convert.ToDouble(oct_cr.Text) + Convert.ToDouble(nov_cr.Text) + Convert.ToDouble(dec_cr.Text) + Convert.ToDouble(jan_cr.Text) + Convert.ToDouble(feb_cr.Text) + Convert.ToDouble(mar_cr.Text));
                final_wd_apr.Text = (Convert.ToDouble(with_apr.Text) + Convert.ToDouble(apr_cr.Text)).ToString();
                final_wd_may.Text = (Convert.ToDouble(with_may.Text) + Convert.ToDouble(may_cr.Text)).ToString();
                final_wd_jun.Text = (Convert.ToDouble(with_jun.Text) + Convert.ToDouble(jun_cr.Text)).ToString();
                final_wd_jul.Text = (Convert.ToDouble(with_jul.Text) + Convert.ToDouble(jul_cr.Text)).ToString();
                final_wd_aug.Text = (Convert.ToDouble(with_aug.Text) + Convert.ToDouble(aug_cr.Text)).ToString();
                final_wd_sep.Text = (Convert.ToDouble(with_sep.Text) + Convert.ToDouble(sep_cr.Text)).ToString();
                final_wd_oct.Text = (Convert.ToDouble(with_oct.Text) + Convert.ToDouble(oct_cr.Text)).ToString();
                final_wd_nov.Text = (Convert.ToDouble(with_nov.Text) + Convert.ToDouble(nov_cr.Text)).ToString();
                final_wd_dec.Text = (Convert.ToDouble(with_dec.Text) + Convert.ToDouble(dec_cr.Text)).ToString();
                final_wd_jan.Text = (Convert.ToDouble(with_jan.Text) + Convert.ToDouble(jan_cr.Text)).ToString();
                final_wd_feb.Text = (Convert.ToDouble(with_feb.Text) + Convert.ToDouble(feb_cr.Text)).ToString();
                final_wd_mar.Text = (Convert.ToDouble(with_mar.Text) + Convert.ToDouble(mar_cr.Text)).ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void opening_bal_KeyPress(object sender, KeyPressEventArgs e)
        {
            char presskey = e.KeyChar;
            if (!char.IsDigit(presskey) && presskey != 8)//8 is backspace
            {
                
                    e.Handled = true;
                
            }

        }

        private void print_Click(object sender, EventArgs e)
        {
            next.Enabled = true;
            try
            {
                string path = @"D:\GPF Repots\" + Name.Text + @"\";
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }
                Document dt = new Document(iTextSharp.text.PageSize.A4.Rotate(), 10, 10, 20, 10);
                PdfWriter wr = PdfWriter.GetInstance(dt, new FileStream(path + year.Text + ".pdf", FileMode.Create));
                dt.Open();
                BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1257, false);
                iTextSharp.text.Font t1 = new iTextSharp.text.Font(bfTimes, 15, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
                iTextSharp.text.Font t2 = new iTextSharp.text.Font(bfTimes, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
                iTextSharp.text.Font t3 = new iTextSharp.text.Font(bfTimes, 11, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                iTextSharp.text.Font t4 = new iTextSharp.text.Font(bfTimes, 8, iTextSharp.text.Font.ITALIC, BaseColor.BLACK);
                iTextSharp.text.Font t5 = new iTextSharp.text.Font(bfTimes, 8, iTextSharp.text.Font.ITALIC, BaseColor.BLUE);
                Paragraph p1 = new Paragraph("Annual GPF Report" + Environment.NewLine +school.Text+ Environment.NewLine, t1);

                p1.Alignment = Element.ALIGN_CENTER;
                p1.SpacingAfter = 20;
                string txt = "";
                if (radioButton1.Checked == true)
                {
                    txt = radioButton1.Text;
                }
                if (radioButton2.Checked == true)
                {
                    txt = radioButton2.Text;

                }
                if (radioButton3.Checked == true)
                {
                    txt = radioButton3.Text;

                }


                Paragraph p2 = new Paragraph("Employee Name: " + Name.Text + "              Year:  " + year.Text + "      Interest Type    " + txt +"     Opening Balance:  "+opening_bal.Text, t2);
                p2.Alignment = Element.ALIGN_CENTER;
                p2.SpacingAfter = 15;
                dt.Add(p1);
                dt.Add(p2);
                PdfPTable table1 = new PdfPTable(10);
                table1.WidthPercentage = 90;
                float[] w2 = { 10f, 12f, 11f, 9f, 11f, 9f, 10f, 10f, 12f, 6f };
                table1.SetWidths(w2);
                table1.AddCell(new Phrase("Month:", t2));
                table1.AddCell(new Phrase("Subscription:", t2));
                table1.AddCell(new Phrase("Bonus/Arrier:", t2));
                table1.AddCell(new Phrase("Refund:", t2));
                table1.AddCell(new Phrase("Total Deposit:", t2));
                table1.AddCell(new Phrase("Temp Advace:", t2));
                table1.AddCell(new Phrase("Other Withdrawals:", t2));
                table1.AddCell(new Phrase("Total Withdrawals:", t2));
                table1.AddCell(new Phrase("Closing Balance:", t2));
                table1.AddCell(new Phrase("Rate %:", t2));
                table1.SpacingAfter = table1.CalculateHeights();
                dt.Add(table1);
                PdfPTable ta2 = new PdfPTable(10);// for april
                ta2.WidthPercentage = 90;
                ta2.SetWidths(w2);
                ta2.AddCell(new Phrase("April", t2));
                ta2.AddCell(new Phrase(Apr.Text, t3));
                ta2.AddCell(new Phrase(apr_b.Text, t3));
                ta2.AddCell(new Phrase(apr_r.Text, t3));
                ta2.AddCell(new Phrase(APR_FD.Text, t3));
                ta2.AddCell(new Phrase(apr_cr.Text, t3));
                ta2.AddCell(new Phrase(with_apr.Text, t3));
                ta2.AddCell(new Phrase(final_wd_apr.Text, t3));
                ta2.AddCell(new Phrase(close_apr.Text, t3));
                ta2.AddCell(new Phrase(m1.Text, t3));
                ta2.SpacingAfter = ta2.CalculateHeights();
                dt.Add(ta2);
                PdfPTable ta3 = new PdfPTable(10);//for may
                ta3.WidthPercentage = 90;
                ta3.SetWidths(w2);
                ta3.AddCell(new Phrase("May", t2));
                ta3.AddCell(new Phrase(May.Text, t3));
                ta3.AddCell(new Phrase(may_b.Text, t3));
                ta3.AddCell(new Phrase(may_r.Text, t3));
                ta3.AddCell(new Phrase(MAY_FD.Text, t3));
                ta3.AddCell(new Phrase(may_cr.Text, t3));
                ta3.AddCell(new Phrase(with_may.Text, t3));
                ta3.AddCell(new Phrase(final_wd_may.Text, t3));
                ta3.AddCell(new Phrase(close_may.Text, t3));
                ta3.AddCell(new Phrase(m2.Text, t3));
                ta3.SpacingAfter = ta3.CalculateHeights();
                dt.Add(ta3);
                PdfPTable ta4 = new PdfPTable(10);//for June
                ta4.WidthPercentage = 90;
                ta4.SetWidths(w2);
                ta4.AddCell(new Phrase("June", t2));
                ta4.AddCell(new Phrase(Jun.Text, t3));
                ta4.AddCell(new Phrase(jun_b.Text, t3));
                ta4.AddCell(new Phrase(jun_r.Text, t3));
                ta4.AddCell(new Phrase(JUN_FD.Text, t3));
                ta4.AddCell(new Phrase(jun_cr.Text, t3));
                ta4.AddCell(new Phrase(with_jun.Text, t3));
                ta4.AddCell(new Phrase(final_wd_jun.Text, t3));
                ta4.AddCell(new Phrase(close_jun.Text, t3));
                ta4.AddCell(new Phrase(m3.Text, t3));
                ta4.SpacingAfter = ta4.CalculateHeights();
                dt.Add(ta4);
                PdfPTable ta5 = new PdfPTable(10);//for july
                ta5.WidthPercentage = 90;
                ta5.SetWidths(w2);
                ta5.AddCell(new Phrase("July", t2));
                ta5.AddCell(new Phrase(Jul.Text, t3));
                ta5.AddCell(new Phrase(jul_b.Text, t3));
                ta5.AddCell(new Phrase(jul_r.Text, t3));
                ta5.AddCell(new Phrase(JUL_FD.Text, t3));
                ta5.AddCell(new Phrase(jul_cr.Text, t3));
                ta5.AddCell(new Phrase(with_jul.Text, t3));
                ta5.AddCell(new Phrase(final_wd_jul.Text, t3));
                ta5.AddCell(new Phrase(close_jul.Text, t3));
                ta5.AddCell(new Phrase(m4.Text, t3));
                ta5.SpacingAfter = ta5.CalculateHeights();
                dt.Add(ta5);
                PdfPTable ta7 = new PdfPTable(10);//for Auguest
                ta7.WidthPercentage = 90;
                ta7.SetWidths(w2);
                ta7.AddCell(new Phrase("August", t2));
                ta7.AddCell(new Phrase(Aug.Text, t3));
                ta7.AddCell(new Phrase(aug_b.Text, t3));
                ta7.AddCell(new Phrase(aug_r.Text, t3));
                ta7.AddCell(new Phrase(AUG_FD.Text, t3));
                ta7.AddCell(new Phrase(aug_cr.Text, t3));
                ta7.AddCell(new Phrase(with_aug.Text, t3));
                ta7.AddCell(new Phrase(final_wd_aug.Text, t3));
                ta7.AddCell(new Phrase(close_aug.Text, t3));
                ta7.AddCell(new Phrase(m5.Text, t3));
                ta7.SpacingAfter = ta7.CalculateHeights();
                dt.Add(ta7);
                PdfPTable ta8 = new PdfPTable(10);//for September
                ta8.WidthPercentage = 90;
                ta8.SetWidths(w2);
                ta8.AddCell(new Phrase("September", t2));
                ta8.AddCell(new Phrase(Sep.Text, t3));
                ta8.AddCell(new Phrase(sep_b.Text, t3));
                ta8.AddCell(new Phrase(sep_r.Text, t3));
                ta8.AddCell(new Phrase(SEP_FD.Text, t3));
                ta8.AddCell(new Phrase(sep_cr.Text, t3));
                ta8.AddCell(new Phrase(with_sep.Text, t3));
                ta8.AddCell(new Phrase(final_wd_sep.Text, t3));
                ta8.AddCell(new Phrase(close_sep.Text, t3));
                ta8.AddCell(new Phrase(m6.Text, t3));
                ta8.SpacingAfter = ta8.CalculateHeights();
                dt.Add(ta8);
                PdfPTable ta9 = new PdfPTable(10);//for October
                ta9.WidthPercentage = 90;
                ta9.SetWidths(w2);
                ta9.AddCell(new Phrase("October", t2));
                ta9.AddCell(new Phrase(Oct.Text, t3));
                ta9.AddCell(new Phrase(oct_b.Text, t3));
                ta9.AddCell(new Phrase(oct_r.Text, t3));
                ta9.AddCell(new Phrase(OCT_FD.Text, t3));
                ta9.AddCell(new Phrase(oct_cr.Text, t3));
                ta9.AddCell(new Phrase(with_oct.Text, t3));
                ta9.AddCell(new Phrase(final_wd_oct.Text, t3));
                ta9.AddCell(new Phrase(close_oct.Text, t3));
                ta9.AddCell(new Phrase(m7.Text, t3));
                ta9.SpacingAfter = ta9.CalculateHeights();
                dt.Add(ta9);
                PdfPTable ta10 = new PdfPTable(10);//for November
                ta10.WidthPercentage = 90;
                ta10.SetWidths(w2);
                ta10.AddCell(new Phrase("November", t2));
                ta10.AddCell(new Phrase(Nov.Text, t3));
                ta10.AddCell(new Phrase(nov_b.Text, t3));
                ta10.AddCell(new Phrase(nov_r.Text, t3));
                ta10.AddCell(new Phrase(NOV_FD.Text, t3));
                ta10.AddCell(new Phrase(nov_cr.Text, t3));
                ta10.AddCell(new Phrase(with_nov.Text, t3));
                ta10.AddCell(new Phrase(final_wd_nov.Text, t3));
                ta10.AddCell(new Phrase(close_nov.Text, t3));
                ta10.AddCell(new Phrase(m8.Text, t3));
                ta10.SpacingAfter = ta10.CalculateHeights();
                dt.Add(ta10);
                PdfPTable ta11 = new PdfPTable(10);//for December
                ta11.WidthPercentage = 90;
                ta11.SetWidths(w2);
                ta11.AddCell(new Phrase("December", t2));
                ta11.AddCell(new Phrase(Dec.Text, t3));
                ta11.AddCell(new Phrase(dec_b.Text, t3));
                ta11.AddCell(new Phrase(dec_r.Text, t3));
                ta11.AddCell(new Phrase(DEC_FD.Text, t3));
                ta11.AddCell(new Phrase(dec_cr.Text, t3));
                ta11.AddCell(new Phrase(with_dec.Text, t3));
                ta11.AddCell(new Phrase(final_wd_dec.Text, t3));
                ta11.AddCell(new Phrase(close_dec.Text, t3));
                ta11.AddCell(new Phrase(m9.Text, t3));
                ta11.SpacingAfter = ta11.CalculateHeights();
                dt.Add(ta11);
                PdfPTable ta12 = new PdfPTable(10);//for January
                ta12.WidthPercentage = 90;
                ta12.SetWidths(w2);
                ta12.AddCell(new Phrase("January", t2));
                ta12.AddCell(new Phrase(Jan.Text, t3));
                ta12.AddCell(new Phrase(jan_b.Text, t3));
                ta12.AddCell(new Phrase(jan_r.Text, t3));
                ta12.AddCell(new Phrase(JAN_FD.Text, t3));
                ta12.AddCell(new Phrase(jan_cr.Text, t3));
                ta12.AddCell(new Phrase(with_jan.Text, t3));
                ta12.AddCell(new Phrase(final_wd_jan.Text, t3));
                ta12.AddCell(new Phrase(close_jan.Text, t3));
                ta12.AddCell(new Phrase(m10.Text, t3));
                ta12.SpacingAfter = ta12.CalculateHeights();
                dt.Add(ta12);
                PdfPTable ta13 = new PdfPTable(10);//for feburary
                ta13.WidthPercentage = 90;
                ta13.SetWidths(w2);
                ta13.AddCell(new Phrase("Feburary", t2));
                ta13.AddCell(new Phrase(Feb.Text, t3));
                ta13.AddCell(new Phrase(feb_b.Text, t3));
                ta13.AddCell(new Phrase(feb_r.Text, t3));
                ta13.AddCell(new Phrase(FEB_FD.Text, t3));
                ta13.AddCell(new Phrase(feb_cr.Text, t3));
                ta13.AddCell(new Phrase(with_feb.Text, t3));
                ta13.AddCell(new Phrase(final_wd_feb.Text, t3));
                ta13.AddCell(new Phrase(close_feb.Text, t3));
                ta13.AddCell(new Phrase(m11.Text, t3));
                ta13.SpacingAfter = ta13.CalculateHeights();
                dt.Add(ta13);
                PdfPTable ta14 = new PdfPTable(10);//for March
                ta14.WidthPercentage = 90;
                ta14.SetWidths(w2);
                ta14.AddCell(new Phrase("March", t2));
                ta14.AddCell(new Phrase(Mar.Text, t3));
                ta14.AddCell(new Phrase(mar_b.Text, t3));
                ta14.AddCell(new Phrase(mar_r.Text, t3));
                ta14.AddCell(new Phrase(MAR_FD.Text, t3));
                ta14.AddCell(new Phrase(mar_cr.Text, t3));
                ta14.AddCell(new Phrase(with_mar.Text, t3));
                ta14.AddCell(new Phrase(final_wd_mar.Text, t3));
                ta14.AddCell(new Phrase(close_mar.Text, t3));
                ta14.AddCell(new Phrase(m12.Text, t3));
                ta14.SpacingAfter = ta14.CalculateHeights();
                dt.Add(ta14);
                PdfPTable sum = new PdfPTable(10);//for G
                sum.WidthPercentage = 90;

                sum.SetWidths(w2);
                sum.AddCell(new Phrase("Grand Total", t2));
                sum.AddCell(new Phrase(sum_subcr.Text, t3));
                sum.AddCell(new Phrase(sum_bonus.Text, t3));
                sum.AddCell(new Phrase(sum_refund.Text, t3));
                sum.AddCell(new Phrase(Sum_Fdep.Text, t3));
                sum.AddCell(new Phrase(sum_tempad.Text, t3));
                sum.AddCell(new Phrase(sum_owith.Text, t3));
                sum.AddCell(new Phrase(sum_fwith.Text, t3));
                sum.AddCell(new Phrase(sum_closing.Text, t3));
                sum.AddCell(new Phrase("", t3));
                sum.SpacingAfter = ta14.CalculateHeights();
                dt.Add(sum);
                Paragraph p3 = new Paragraph(l28.Text + "   " + intt.Text + "               Total Withdrawals   " + sum_fwith.Text + Environment.NewLine + l26.Text + "  " + closing_bal.Text + "    " + l29.Text + "    " + label4.Text + Environment.NewLine + "( " + label5.Text + " )", t3);
                p3.Alignment = Element.ALIGN_CENTER;
                dt.Add(p3);
                p3.SpacingAfter = 20;
                Paragraph p4 = new Paragraph(Environment.NewLine + Environment.NewLine + "Prepared By:" + User.Text + "                                           Checked By:                                                 Verified By:                                                   Principal Stamp" + Environment.NewLine + Environment.NewLine + Environment.NewLine + Environment.NewLine + Environment.NewLine, t3);
                dt.Add(p4);
                p4.SpacingAfter = 50;
                Paragraph p5 = new Paragraph("** This document is invalid without Principal's Stamp & Signature", t4);
                p5.Alignment = Element.CREATOR;
                dt.Add(p5);
                Paragraph p6 = new Paragraph("This Software is Design & Developed By Bijalwan Enterprises Contact For Help/Query &  More 7060629794 Naveenprasadbijlwan@gmail.com", t5);
                p6.Alignment = Element.ALIGN_BASELINE;
                dt.Add(p6);
                dt.Close();
                MessageBox.Show("File Saved Successfully.!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void Name_KeyPress(object sender, KeyPressEventArgs e)
        {
            char presskey = e.KeyChar;
            if (char.IsDigit(presskey))
            {

                e.Handled = true;

            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", @"D:\GPF Repots\" + Name.Text + @"\");
        }
        private void year_Leave(object sender, EventArgs e)
        {
            MessageBox.Show("Interest Fetched");    
            try
            {
                int yr = Convert.ToInt32(year.Text.ToString().Substring(0, 4));
                if (yr > 1979 && yr < 1990)
                {
                    if (yr == 1980)
                    {
                        radioButton1.Checked = true;
                        annual.Text = "8.50";
                    }
                    else if (yr == 1981)
                    {


                        radioButton1.Checked = true; annual.Text = "8.50";
                    }
                    else if (yr == 1982)
                    {


                        radioButton1.Checked = true; annual.Text = "9.00";
                    }
                    else if (yr == 1983)
                    {


                        radioButton1.Checked = true; annual.Text = "9.50";
                    }
                    else if (yr == 1984)
                    {


                        radioButton1.Checked = true; annual.Text = "10.00";
                    }
                    else if (yr == 1985)
                    {


                        radioButton1.Checked = true; annual.Text = "10.05";
                    }
                    else if (yr == 1986)
                    {

                        radioButton1.Checked = true; annual.Text = "12.00";

                    }
                    else if (yr == 1987)
                    {


                        radioButton1.Checked = true; annual.Text = "12.00";
                    }
                    else if (yr == 1988)
                    {


                        radioButton1.Checked = true; annual.Text = "12.00";
                    }
                    else if (yr == 1989)
                    {


                        radioButton1.Checked = true; annual.Text = "12.00";
                    }
                    else
                    {
                        MessageBox.Show("Oops.!" + Environment.NewLine + Environment.NewLine + " No intrest rate found.!  you can enter manually");

                    }
                }
                if (yr > 1989 && yr < 2000)
                {
                    if (yr == 1990)
                    {


                        radioButton1.Checked = true; annual.Text = "12.00";
                    }
                    else if (yr == 1991)
                    {


                        radioButton1.Checked = true; annual.Text = "12.00";
                    }
                    else if (yr == 1992)
                    {

                        radioButton1.Checked = true;
                        annual.Text = "12.00";
                    }
                    else if (yr == 1993)
                    {
                        annual.Text = "12.00";

                        radioButton1.Checked = true;
                    }
                    else if (yr == 1994)
                    {

                        radioButton1.Checked = true;
                        annual.Text = "12.00";
                    }
                    else if (yr == 1995)
                    {


                        radioButton1.Checked = true; annual.Text = "12.00";
                    }
                    else if (yr == 1996)
                    {

                        radioButton1.Checked = true;
                        annual.Text = "12.00";
                    }
                    else if (yr == 1997)
                    {

                        radioButton1.Checked = true;
                        annual.Text = "12.00";
                    }
                    else if (yr == 1998)
                    {

                        radioButton1.Checked = true;
                        annual.Text = "12.00";
                    }
                    else if (yr == 1999)
                    {

                        radioButton1.Checked = true;
                        annual.Text = "12.00";
                    }
                    else
                    {
                        MessageBox.Show("Oops.!" + Environment.NewLine + Environment.NewLine + " No intrest rate found.!  you can enter manually");

                    }

                }
                if (yr > 1999 && yr < 2010)
                {
                    if (yr == 2000)
                    {

                        radioButton1.Checked = true;
                        annual.Text = "11.00";
                    }
                    else if (yr == 2001)
                    {


                        radioButton1.Checked = true; annual.Text = "9.05";
                    }
                    else if (yr == 2002)
                    {


                        radioButton1.Checked = true; annual.Text = "9.00";
                    }
                    else if (yr == 2003)
                    {


                        radioButton1.Checked = true; annual.Text = "8.00";
                    }
                    else if (yr == 2004)
                    {


                        radioButton1.Checked = true; annual.Text = "8.00";
                    }
                    else if (yr == 2005)
                    {

                        radioButton1.Checked = true;
                        annual.Text = "8.00";
                    }
                    else if (yr == 2006)
                    {


                        radioButton1.Checked = true; annual.Text = "8.00";
                    }
                    else if (yr == 2007)
                    {

                        radioButton1.Checked = true;
                        annual.Text = "8.00";
                    }
                    else if (yr == 2008)
                    {

                        radioButton1.Checked = true;
                        annual.Text = "8.00";
                    }
                    else if (yr == 2009)
                    {
                        radioButton1.Checked = true;

                        annual.Text = "8.00";
                    }
                    else
                    {
                        MessageBox.Show("Oops.!" + Environment.NewLine + Environment.NewLine + " No intrest rate found.!  you can enter manually");

                    }
                }
                if (yr > 2010 && yr < 2020)
                {
                    if (yr == 2010)
                    {


                        radioButton1.Checked = true; annual.Text = "8.00";
                    }
                    else if (yr == 2011)
                    {
                        radioButton3.Checked = true;
                        m1.Text = m2.Text = m3.Text = m4.Text = m5.Text = m6.Text = m7.Text = m8.Text = "8.00";
                        m9.Text = m10.Text = m11.Text = m12.Text = "8.06";

                    }
                    else if (yr == 2012)
                    {
                        radioButton1.Checked = true; annual.Text = "8.80";

                    }
                    else if (yr == 2013)
                    {
                        radioButton1.Checked = true; annual.Text = "8.70";

                    }
                    else if (yr == 2014)
                    {
                        radioButton1.Checked = true; annual.Text = "8.70";

                    }
                    else if (yr == 2015)
                    {
                        radioButton1.Checked = true; annual.Text = "8.70";

                    }
                    else if (yr == 2016)
                    {
                        radioButton2.Checked = true; ;
                        quater1.Text = quater2.Text = "8.10";
                        quater3.Text = quater4.Text = "8.00";
                    }
                    else if (yr == 2017)
                    {
                        radioButton2.Checked = true;
                        quater1.Text = "7.90";
                        quater2.Text = quater3.Text = "7.80";
                        quater4.Text = "7.60";

                    }
                    else if (yr == 2018)
                    {
                        radioButton2.Checked = true;
                        quater1.Text = quater2.Text = "7.60";
                        quater3.Text = quater4.Text = "8.00";
                    }
                    else if (yr == 2019)
                    {
                        radioButton2.Checked = true;
                        quater1.Text = "8.00";
                        quater2.Text = quater3.Text = quater4.Text = "7.90";


                    }
                    else
                    {
                        MessageBox.Show("Oops.!" + Environment.NewLine + Environment.NewLine + " No intrest rate found.!  you can enter manually");

                    }
                }
                if (radioButton1.Checked == true)
                {
                    m1.Text = m2.Text = m3.Text = m4.Text = m5.Text = m6.Text = m7.Text = m8.Text = m9.Text = m10.Text = m11.Text = m12.Text = annual.Text;
                }
                else if (radioButton2.Checked == true)
                {
                    m1.Text = m2.Text = m3.Text = quater1.Text;
                    m4.Text = m5.Text = m6.Text = quater2.Text;
                    m7.Text = m8.Text = m9.Text = quater3.Text;
                    m10.Text = m11.Text = m12.Text = quater4.Text;

                }
            }
            catch (Exception) {
                MessageBox.Show("Year Can't Null");
            }
        }


        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked == true)
            {
                annual.Visible = false;
                quater1.Visible = false;
                quater2.Visible = false;
                quater3.Visible = false;
                quater4.Visible = false;

            }

        }

        private void next_Click(object sender, EventArgs e)
        {
            string bal = "";
            bal= Math.Round(Convert.ToDouble(closing_bal.Text),0).ToString();            
            opening_bal.Text = bal;

            final_wd_apr.Text = final_wd_may.Text = final_wd_jun.Text = final_wd_jul.Text = final_wd_aug.Text = final_wd_sep.Text = final_wd_oct.Text = final_wd_nov.Text = final_wd_dec.Text = final_wd_jan.Text = final_wd_feb.Text = final_wd_mar.Text = apr_b.Text = apr_cr.Text = with_apr.Text = may_b.Text = may_cr.Text = with_may.Text =  jun_b.Text = jun_cr.Text  = with_jun.Text = jul_b.Text = jul_cr.Text = with_jul.Text =  aug_b.Text = aug_cr.Text = with_aug.Text = sep_b.Text = sep_cr.Text  = with_sep.Text =   oct_b.Text = oct_cr.Text  = with_oct.Text = nov_b.Text = nov_cr.Text = with_nov.Text = dec_b.Text = dec_cr.Text  = with_dec.Text = jan_b.Text = jan_cr.Text = with_jan.Text = feb_b.Text = feb_cr.Text  = with_feb.Text =   mar_b.Text = mar_cr.Text  = with_mar.Text = "0";
            string y = year.Text.ToString().Substring(0, 4);
           int x= Convert.ToInt32(y);
            x++;

            int y1 = x+1;
           year.Text = x.ToString() + "-" + y1.ToString().Substring(2);
            year.Focus();   
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton2.Checked==true)
            {
                annual.Visible = false;
                quater1.Visible = true; quater2.Visible = true; quater3.Visible = true; quater4.Visible = true;

            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                annual.Visible = true;
                quater1.Visible = false; quater2.Visible = false; quater3.Visible = false; quater4.Visible = false;

            }
        }

        private void radioButton2_Leave(object sender, EventArgs e)
        {
            if(radioButton1.Checked==true)
            {
               m1.Text=m2.Text = m3.Text = m4.Text = m5.Text = m6.Text = m7.Text = m8.Text = m9.Text = m10.Text = m11.Text = m12.Text = annual.Text;
            }
           else if(radioButton2.Checked==true)
            {
                m1.Text = m2.Text = m3.Text = quater1.Text;
                m4.Text = m5.Text = m6.Text = quater2.Text;
                m7.Text = m8.Text = m9.Text = quater3.Text;
                m10.Text = m11.Text = m12.Text = quater4.Text;

            }
        }

        private void m1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char presskey = e.KeyChar;
            if (char.IsLetter(presskey))
            {

                e.Handled = true;

            }
        }
    }
}
