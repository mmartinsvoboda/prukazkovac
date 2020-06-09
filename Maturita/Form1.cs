using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Printing;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.Actions;
using PdfSharp.Forms;

namespace Maturita
{
    public partial class Průkazkovač : Form
    {
        public Průkazkovač()
        {
            InitializeComponent();
            Application.ApplicationExit += new EventHandler(this.OnApplicationExit);
        }

        //Variables
        string file;
        string photoPath;
        string fileName2;

        Object filename2;

        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;

        Excel.Application xlApp2;
        Excel.Workbook xlWorkBook2;
        Excel.Worksheet xlWorkSheet2;

        object missing = System.Reflection.Missing.Value;
        Word.Application wordApp = new Word.Application();
        Word.Document aDoc = null;

        int rCnt;
        int rCntExp;

        double pocetStudentu = 0;

        bool loaded = false;
        bool opened = false;
        //Part 1 - Načtení databáze
        /// <summary>
        /// Načtení databáze
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void loadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Vyberte soubor databáze.", "Vyberte soubor databáze",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            DialogResult result = openFileDialog1.ShowDialog();
            if (result != DialogResult.OK) // DialogResult musí být OK
                return;

            file = openFileDialog1.FileName;

            //Loading Excel File...
            toolStripStatusLabel1.Text = "Načítám Excel soubor...";
            try
            {
                loadExcelFile(file);
            }
            catch (Exception Ex)
            {
                exceptionsHandler(Ex);
                return;
            }

            toolStripStatusLabel1.Text = "Vytvářím soubor vlastní databáze...";
            try
            {
                createExcelFile();
            }
            catch (Exception Ex)
            {
                exceptionsHandler(Ex);
                return;
            }

            toolStripStatusLabel1.Text = "Výběr cesty ke složce s obrázky...";
            MessageBox.Show("Vyberte složku s fotografiemi.", "Info",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            try
            {
                photoPath = vyberSlozku();
            }
            catch (Exception Ex)
            {
                exceptionsHandler(Ex);
                return;
            }

            toolStripStatusLabel1.Text = "Vytvářím databázi. Může to trvat i několik minut.";
            try
            {
                fillExcelFile();
            }
            catch (Exception Ex)
            {
                exceptionsHandler(Ex);
                return;
            }

            //Ukončení Excel souboru s databází
            try
            {
                xlWorkBook.Close(false, null, null);
                xlApp.Quit();
            }
            catch (Exception Ex)
            {
                exceptionsHandler(Ex);
                return;
            }

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            //
            toolStripStatusLabel1.Text = "Probíhá kontrola jednotlivých položek...";
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                try
                {
                    checkPhoto(i);
                }
                catch (Exception Ex)
                {
                    exceptionsHandler(Ex);
                }
            }

            toolStripStatusLabel1.Text = "Všechny položky byly úspěšně vloženy!";
            listBox1.Enabled = true;
            button4.Text = "Vytvořit kartičky";
            loaded = true;
            toolStripMenuItem1.Visible = true;
            loadToolStripMenuItem.Visible = false;
        }

        /// <summary>
        /// Načtení Excel souboru databáze
        /// </summary>
        /// <param name="file">Cesta k souboru s databází</param>
        private void loadExcelFile(string file)
        {
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
        }

        /// <summary>
        /// //Tvorba dočasného Excel souboru
        /// </summary>
        private void createExcelFile()
        {
            Excel.Application xlApp2 = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp2 == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


            object misValue = System.Reflection.Missing.Value;

            xlWorkBook2 = xlApp2.Workbooks.Add(misValue);
            xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1);
        }

        /// <summary>
        /// Vyplnění dočasného excel souboru
        /// </summary>
        private void fillExcelFile()
        {
            xlWorkSheet2.Cells[1, 1] = "Jméno";
            xlWorkSheet2.Cells[1, 2] = "Příjmení";
            xlWorkSheet2.Cells[1, 3] = "Datum narození";
            xlWorkSheet2.Cells[1, 4] = "platnost do";
            xlWorkSheet2.Cells[1, 5] = "ID";
            xlWorkSheet2.Cells[1, 6] = "Cesta k fotografii";
            xlWorkSheet2.Cells[1, 7] = "Název fotky";
            xlWorkSheet2.Cells[1, 8] = "Správnost názvu";
            xlWorkSheet2.Cells[1, 9] = "Checked";
            xlWorkSheet2.Cells[1, 10] = "Třída";

            range = xlWorkSheet.UsedRange;

            for (rCnt = 2; rCnt <= range.Rows.Count + 1; rCnt++)
            {
                if (xlWorkSheet.Cells[rCnt, 2].Text == "")
                {
                    Console.WriteLine(pocetStudentu);
                    break;
                }

                //xlWorkSheet.Cells[1, 1] = "Sheet 1 content";
                //Vyplň jméno
                xlWorkSheet2.Cells[rCnt, 1] = xlWorkSheet.Cells[rCnt, 5];

                //Vyplň příjmení
                xlWorkSheet2.Cells[rCnt, 2] = xlWorkSheet.Cells[rCnt, 4];

                //Vyplň Datum narození
                xlWorkSheet2.Cells[rCnt, 3] = xlWorkSheet.Cells[rCnt, 8];

                //Vyplň Platnost
                xlWorkSheet2.Cells[rCnt, 4] = xlWorkSheet.Cells[rCnt, 13];

                //Vyplň ID
                xlWorkSheet2.Cells[rCnt, 5] = xlWorkSheet.Cells[rCnt, 2];

                //Vyplň cestu k fotografii
                string fotka = xlWorkSheet.Cells[rCnt, 10].Text.ToLower() + "_" +
                    (RemoveDiacritics(xlWorkSheet.Cells[rCnt, 4].Text)).ToLower() + "_" +
                    (RemoveDiacritics(xlWorkSheet.Cells[rCnt, 5].Text).ToLower()) + ".jpg";
                string cesta;
                cesta = photoPath + "/" + fotka;
                xlWorkSheet2.Cells[rCnt, 6] = cesta;
                xlWorkSheet2.Cells[rCnt, 7] = fotka;
                xlWorkSheet2.Cells[rCnt, 8] = "1";
                xlWorkSheet2.Cells[rCnt, 9] = "0";

                Console.WriteLine();
                Console.WriteLine(cesta);
                Console.WriteLine();

                //Vyplň třídu
                xlWorkSheet2.Cells[rCnt, 10] = xlWorkSheet.Cells[rCnt, 10];

                //Vyplň ListBox1
                fillListBox1(rCnt);

                pocetStudentu++;
            }
        }

        /// <summary>
        /// Vyplňění ListBox1
        /// </summary>
        /// <param name="radek"></param>
        private void fillListBox1(int radek)
        {
            try
            {
                listBox1.Items.Add(xlWorkSheet2.Cells[radek, 1].Text + " " + xlWorkSheet2.Cells[radek, 2].Text);
            }
            catch (Exception Ex)
            {
                exceptionsHandler(Ex);
            }
        }
        //Konec Part1

        //Part 2 - Úprava studentů
        /// <summary>
        /// Event - Spouští se při změně položky v ListBoxu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "Zobrazuji studenta " + xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1].Text + " "
                + xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2].Text;

            button7.Visible = true;

            label1.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1].Text + " "
                + xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2].Text;
            label2.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 3].Text;
            label3.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 4].Text;
            label4.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 5].Text;
            label6.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 7].Text;

            if (!System.IO.File.Exists(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 6].Text))
            {
                label6.ForeColor = System.Drawing.Color.Red;
                button5.Visible = true;
                button6.Visible = true;
            }
            else
            {
                label6.ForeColor = System.Drawing.Color.Black;
                button5.Visible = false;
                button6.Visible = false;
                //HAVE NOT BEEN TESTED YET
                xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 9] = "1";
                pictureBox2.Image = Image.FromFile(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 6].Text);

            }

            if (xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 9].Text == "1")
            {
                if (System.IO.File.Exists(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 6].Text))
                {
                    Console.WriteLine("File {0} DOES exist.", xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 6].Text);
                    pictureBox2.Image = Image.FromFile(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 6].Text);
                    pictureBox2.SizeMode = PictureBoxSizeMode.StretchImage;
                    Console.WriteLine("Shoud have changed!");
                }
                else
                {
                    Console.WriteLine("File {0} DOESN'T EXIST.", xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 6].Text);
                    pictureBox2.Image = pictureBox2.ErrorImage;
                    pictureBox2.SizeMode = PictureBoxSizeMode.Normal;
                    MessageBox.Show("Název souboru fotografie není platný!", "Chyba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 9] = "0";
                    listBox1_SelectedIndexChanged(this, null);
                }
            }
            else
            {
                pictureBox2.Image = pictureBox2.ErrorImage;
                pictureBox2.SizeMode = PictureBoxSizeMode.Normal;
            }

            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            label4.Visible = true;

            button1.Visible = true;
        }

        /// <summary>
        /// Potvrzení názvu fotografie, který byl označen jako chybný
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            button5.Visible = false;
            string fotka = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 10].Text.ToLower() + "_" +
                    (RemoveDiacritics(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2].Text)).ToLower() + "_" +
                    (RemoveDiacritics(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1].Text).ToLower()) + ".jpg";
            string cesta;
            cesta = photoPath + "/" + fotka;
            xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 6] = cesta;
            xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 7] = fotka;

            if (System.IO.File.Exists(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 6].Text))
            {
                listBox1.Items[listBox1.SelectedIndex] = "EDITED: " + xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1].Text + " "
        + xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2].Text;
                xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 9] = "1";
            }
            else
            {
                MessageBox.Show("Název souboru fotografie není platný!", "Chyba",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            listBox1_SelectedIndexChanged(this, null);
        }

        /// <summary>
        /// Název fotografie bude znovu vytvořen ze jména studenta ( V případě změny údajů v aplikaci )
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Opravdu chcete aktualizovat název fotografie?", "Potvrzení aktualizace",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 7] =
                    xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 10].Text.ToLower() + "_" +
                    (RemoveDiacritics(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2].Text)).ToLower() + "_" +
                    (RemoveDiacritics(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1].Text).ToLower()) + ".jpg";

                listBox1_SelectedIndexChanged(this, null);
            }
        }

        /// <summary>
        /// Vyhledání nové cesty k fotografii
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            openFileDialog3.ShowDialog();
            string result = openFileDialog3.FileName;
            string result2 = openFileDialog3.SafeFileName;

            xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 6] = result;
            xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 7] = result2;

            checkPhoto(listBox1.SelectedIndex);
            listBox1_SelectedIndexChanged(this, null);
        }

        //Úprava studentů
        /// <summary>
        /// UI se přepne do verze pro úpravu studenta
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Enabled = false;
            button1.Visible = false;
            button2.Visible = true;
            button3.Visible = true;

            textBox1.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1].Text;
            textBox5.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2].Text;
            textBox2.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 3].Text;
            textBox3.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 4].Text;
            textBox4.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 5].Text;

            textBox1.Visible = true;
            textBox2.Visible = true;
            textBox3.Visible = true;
            textBox4.Visible = true;
            textBox5.Visible = true;

            label1.Visible = false;
            label2.Visible = false;
            label3.Visible = false;
            label4.Visible = false;
        }

        /// <summary>
        /// Uložení změn provedených při Úpravách studenta, dotaz na aktualizaci názvu fotografie
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1] = textBox1.Text;
            xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2] = textBox5.Text;
            xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 3] = textBox2.Text;
            xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 4] = textBox3.Text;
            xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 5] = textBox4.Text;

            if (Convert.ToInt32(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 9].Text) == 1)
            {
                listBox1.Items[listBox1.SelectedIndex] = (xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1].Text
                    + " " + xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2].Text);
            }


            if (Convert.ToInt32(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 9].Text) == 0)
            {
                if (MessageBox.Show("Chcete na základě provedených změn aktualizovat název fotografie?", "Změna názvu fotky",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string fotka = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 10].Text.ToLower() + "_" +
                        (RemoveDiacritics(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2].Text)).ToLower() + "_" +
                        (RemoveDiacritics(xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1].Text).ToLower()) + ".jpg";
                    string cesta;
                    cesta = photoPath + "/" + fotka;
                    xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 6] = cesta;
                    xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 7] = fotka;

                    xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 9] = "1";

                    listBox1.Items[listBox1.SelectedIndex] = ("EDITED: " +
                        xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1].Text + " "
            + xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2].Text);

                }
            }

            textBox1.Visible = false;
            textBox2.Visible = false;
            textBox3.Visible = false;
            textBox4.Visible = false;
            textBox5.Visible = false;

            label1.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1].Text + " "
                + xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2].Text;
            label2.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 3].Text;
            label3.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 4].Text;
            label4.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 5].Text;

            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            label4.Visible = true;

            listBox1.Enabled = true;
            button1.Visible = true;
            button2.Visible = false;
            button3.Visible = false;

            listBox1_SelectedIndexChanged(this, null);
        }

        /// <summary>
        /// Zrušení změn informací o studentovi, informace nebudou uloženy
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Visible = false;
            textBox2.Visible = false;
            textBox3.Visible = false;
            textBox4.Visible = false;
            textBox5.Visible = false;

            label1.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 1].Text + " "
                + xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 2].Text;
            label2.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 3].Text;
            label3.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 4].Text;
            label4.Text = xlWorkSheet2.Cells[listBox1.SelectedIndex + 2, 5].Text;

            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            label4.Visible = true;

            listBox1.Enabled = true;
            button1.Visible = true;
            button2.Visible = false;
            button3.Visible = false;
        }
        //Konec Part 2

        //Part 3 - Tvorba samotných kartiček
        //Ověřovací metody
        /// <summary>
        /// Button - V případě, že je vše jak má, spustí tvorbu kartiček
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            if (loaded)
            {
                wordOutput();
            }
            else
            {
                loadToolStripMenuItem_Click(this, null);
            }

        }

        /// <summary>
        /// Tlačítko ve StripMenu - V případě, že je vše jak má, spustí tvorbu kartiček
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (loaded)
                wordOutput();
        }

        /// <summary>
        /// Ověření, zda je vybraný Word platnou šablonou pro 16 studentů
        /// </summary>
        /// <returns>False - Aplikace nebude pokračovat, uživatel musí zvolit jinou šablonu</returns>
        private bool checkWord()
        {
            bool bezChyby = true;
            try
            {
                rCntExp = 2;
                fillWord();
            }
            catch //(Exception e)
            {
                bezChyby = false;
            }

            aDoc.Close(false, ref missing, ref missing);
            return bezChyby;
        }

        //Vyplnění Template, vytvoření PDF souborů, následné spojení
        /// <summary>
        /// Základní metoda, ve které jsou volány následující metody pro vytvoření výsledného PDF souboru
        /// </summary>
        private void wordOutput()
        {
            if (!finalCheck())
            {
                MessageBox.Show("Některé položky obsahují neplatnou cestu k fotografii.", "Chyba",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                CreateWordDocument();
            }
            catch (Exception Ex)
            {
                exceptionsHandler(Ex);
                return;
            }

            if (!opened)
            {
                toolStripStatusLabel1.Text = "Načtěte, prosím, šablonu.";
                return;
            }

            rCntExp = 2;

            if (!checkWord())
            {
                MessageBox.Show("Vybraný dokument není platná šablona.", "Neplatná šablona",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            toolStripStatusLabel1.Text = "Probíhá Export PDF souboru, to může trvat několik minut...";

            try
            {
                aDoc = wordApp.Documents.Open(ref filename2, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

                aDoc.Activate();
            }
            catch (Exception Ex)
            {
                exceptionsHandler(Ex);
                return;
            }

            rCntExp = 2;

            for (int i = 0; i < pocitejStranky(); i++)
            {
                try
                {
                    fillWord();
                }
                catch (Exception Ex)
                {
                    exceptionsHandler(Ex);
                    return;
                }

                object pdfName = (AppDomain.CurrentDomain.BaseDirectory + Convert.ToString(i) + ".pdf");

                //Save
                object fileFormat = Word.WdSaveFormat.wdFormatPDF;

                // Save document into PDF Format
                try
                {
                    aDoc.SaveAs(ref pdfName,
                        ref fileFormat, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing);


                    aDoc.Close(false, ref missing, ref missing);

                    aDoc = wordApp.Documents.Open(ref filename2, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing);
                }
                catch (Exception Ex)
                {
                    exceptionsHandler(Ex);
                    return;
                }

                Console.WriteLine(fileFormat);
            }

            aDoc.Close(false, ref missing, ref missing);
            wordApp.Quit();


            spojeniPDF();


            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;

            loadToolStripMenuItem.Enabled = false;
            toolStripMenuItem1.Enabled = false;

            label7.Visible = true;
            label8.Text = AppDomain.CurrentDomain.BaseDirectory + "Karticky.pdf";
            label8.Visible = true;
            button8.Visible = true;

            toolStripStatusLabel1.Text = "Váš soubor byl úspěšně vytvořen...";
            dotazOtevreni();
        }

        /// <summary>
        /// Otevření template Word souboru
        /// </summary>
        private void CreateWordDocument()
        {
            toolStripStatusLabel1.Text = "Vytvářím dokument Word...";

            try
            {
                wordApp.Visible = false;
            }
            catch (Exception Ex)
            {
                exceptionsHandler(Ex);
                return;
            }

            MessageBox.Show("Vyberte šablonu.", "Vyberte šablonu", MessageBoxButtons.OK, MessageBoxIcon.Information);

            filename2 = (Object)vyberSoubor();

            if (!opened)
                return;

            try
            {
                aDoc = wordApp.Documents.Open(ref filename2, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);
                aDoc.Activate();
            }
            catch (Exception Ex)
            {
                exceptionsHandler(Ex);
                return;
            }
        }

        /// <summary>
        /// Vyplnění položek ve Word souboru
        /// </summary>
        private void fillWord()
        {
            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name1"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date1"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid1"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id1"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            var shape = aDoc.Bookmarks["photo1"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name2"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date2"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid2"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id2"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo2"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }


            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name3"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date3"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid3"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id3"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo3"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name4"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date4"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid4"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id4"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo4"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "" || xlWorkSheet2.Cells[rCntExp, 1].Text == null)
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name5"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date5"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid5"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id5"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo5"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name6"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date6"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid6"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id6"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo6"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name7"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date7"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid7"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id7"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo7"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name8"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date8"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid8"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id8"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo8"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name9"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date9"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid9"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id9"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo9"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name10"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date10"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid10"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id10"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo10"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name11"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date11"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid11"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id11"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo11"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name12"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date12"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid12"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id12"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo12"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name13"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date13"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid13"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id13"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo13"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name14"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date14"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid14"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id14"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo14"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name15"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date15"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid15"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id15"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo15"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
            if (xlWorkSheet2.Cells[rCntExp, 1].Text == "")
            {
                return;
            }

            Console.WriteLine("Filling Name...");
            aDoc.Bookmarks["name16"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 1].Text + " "
                + xlWorkSheet2.Cells[rCntExp, 2].Text);
            Console.WriteLine("Name filled!");

            Console.WriteLine("Filling Date...");
            aDoc.Bookmarks["date16"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 3].Text);
            Console.WriteLine("Date filled!");

            Console.WriteLine("Filling Valid...");
            aDoc.Bookmarks["valid16"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 4].Text);
            Console.WriteLine("Valid filled!");

            Console.WriteLine("Filling ID...");
            aDoc.Bookmarks["id16"].Select();
            wordApp.Selection.TypeText(xlWorkSheet2.Cells[rCntExp, 5].Text);
            Console.WriteLine("ID filled!");

            Console.WriteLine("Insreting Photo...");
            shape = aDoc.Bookmarks["photo16"].Range.InlineShapes.AddPicture
                (xlWorkSheet2.Cells[rCntExp, 6].Text, false, true);
            shape.Width = 51.15F;
            shape.Height = 68.45F;
            Console.WriteLine("Photo inserted!");

            rCntExp++;
        }

        /// <summary>
        /// Spojení vytvořených PDF souborů do výsledného PDF souboru, včetně jejich smazání
        /// </summary>
        private void spojeniPDF()
        {
            // Create the output document
            PdfDocument outputDocument = new PdfDocument();

            PdfDocument inputDocument;
            for (int i = 0; i < pocitejStranky(); i++)
            {
                fileName2 = (AppDomain.CurrentDomain.BaseDirectory + Convert.ToString(i) + ".pdf");

                // Open the document to import pages from it.
                inputDocument = PdfReader.Open(fileName2, PdfDocumentOpenMode.Import);

                // Iterate pages
                int count = inputDocument.PageCount;
                for (int idx = 0; idx < count; idx++)
                {
                    // Get the page from the external document...
                    PdfPage page = inputDocument.Pages[idx];
                    // ...and add it to the output document.
                    outputDocument.AddPage(page);
                }
            }

            // Save the document...
            const string filename = "Karticky.pdf";
            outputDocument.Save(filename);

            for (int i = 0; i < pocitejStranky(); i++)
            {
                fileName2 = (AppDomain.CurrentDomain.BaseDirectory + Convert.ToString(i) + ".pdf");
                if (System.IO.File.Exists(fileName2))
                {
                    // Use a try block to catch IOExceptions, to
                    // handle the case of the file already being
                    // opened by another process.
                    try
                    {
                        System.IO.File.Delete(fileName2);
                    }
                    catch (System.IO.IOException e)
                    {
                        Console.WriteLine(e.Message);
                        return;
                    }
                }
            }
        }

        //Otevření výsledného PDF souboru
        /// <summary>
        /// MesssageBox - Dotaz na otevření výsledného PDF souboru, Otevření PDF souboru
        /// </summary>
        private void dotazOtevreni()
        {
            if (MessageBox.Show("Soubor byl úspěšně vytvořen ve složce s aplikací.\n\n"
                + AppDomain.CurrentDomain.BaseDirectory + "Karticky.pdf\n\n" + "Přejete si jej zobrazit?", "Hotovo",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    System.Diagnostics.Process.Start(AppDomain.CurrentDomain.BaseDirectory + "/Karticky.pdf");
                }
                catch (Exception Ex)
                {
                    exceptionsHandler(Ex);
                    return;
                }
            }
        }

        /// <summary>
        /// Otevření výsledného PDF souboru přes Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(AppDomain.CurrentDomain.BaseDirectory + "Karticky.pdf");
            }
            catch (Exception Ex)
            {
                exceptionsHandler(Ex);
                return;
            }
        }


        /// <summary>
        /// Výpočet, kolik stránek bude potřeba vyplnit v závislosti na počtu položek
        /// </summary>
        /// <returns></returns>
        private int pocitejStranky()
        {
            double a;
            a = pocetStudentu / 16;
            double b;
            b = a % 1;
            if (b >= 0.5f)
            {
                a = Convert.ToInt32(a);
            }
            else
            {
                a = Convert.ToInt32(a + 1);
            }

            return Convert.ToInt32(a);
        }
        //Konec Part 3

        // Part 4 - Ostatní metody
        /// <summary>
        /// ToolStripMenuItem - Zobrazí informace o aplikaci
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Autor: Martin Svoboda\nProgram byl vytvořen 7. 4. 2016 jako maturitní práce.\n\nE-mail: martinsvoboda@atlas.cz",
                "Informace");
        }

        /// <summary>
        /// Kontrola názvu fotky pro přítomnost nevyžádaných znaků (mezer) a ověření existence souboru
        /// </summary>
        /// <param name="cislo"></param>
        private void checkPhoto(int cislo)
        {

            if (Convert.ToInt32(xlWorkSheet2.Cells[cislo + 2, 9].Text) == 0)
            {
                string fotka = xlWorkSheet2.Cells[cislo + 2, 7].Text;
                foreach (char znak in fotka)
                {
                    if (znak == 32)
                    {
                        listBox1.Items[cislo] = "    !!!    " + xlWorkSheet2.Cells[cislo + 2, 1].Text + " "
                            + xlWorkSheet2.Cells[cislo + 2, 2].Text + "     !!!     ";
                        xlWorkSheet2.Cells[cislo + 2, 8] = "0";
                        xlWorkSheet2.Cells[cislo + 2, 9] = "0";
                        break;
                    }
                    else
                    {
                        xlWorkSheet2.Cells[cislo + 2, 8] = "1";
                        xlWorkSheet2.Cells[cislo + 2, 9] = "1";
                    }

                }

            }

            if (System.IO.File.Exists(xlWorkSheet2.Cells[cislo + 2, 6].Text))
            {
                xlWorkSheet2.Cells[cislo + 2, 8] = "1";
                xlWorkSheet2.Cells[cislo + 2, 9] = "1";
                listBox1.Items[cislo] = xlWorkSheet2.Cells[cislo + 2, 1].Text + " " + xlWorkSheet2.Cells[cislo + 2, 2].Text;
            }
            else
            {
                listBox1.Items[cislo] = "    !!!    " + xlWorkSheet2.Cells[cislo + 2, 1].Text + " "
                    + xlWorkSheet2.Cells[cislo + 2, 2].Text + "     !!!     ";
                xlWorkSheet2.Cells[cislo + 2, 8] = "0";
                xlWorkSheet2.Cells[cislo + 2, 9] = "0";
            }

        }

        /// <summary>
        /// Kontrola, jestli všechny fotografie opravdu existují. 
        /// Chybné fotografie jsou označeny v ListBox1
        /// Dokud nebudou všechny fotografie nalezeny, program nedovolí vytvořit kartičky
        /// </summary>
        /// <returns>True - Program může pokračovat ve tvorbě kartiček</returns>
        private bool finalCheck()
        {

            bool plati = true;

            for (int i = 2; i < i + 1; i++)
            {
                if (System.IO.File.Exists(xlWorkSheet2.Cells[i, 6].Text))
                {
                    //Pokračuje se v kontrole další položky
                }
                else
                {
                    listBox1.Items[i - 2] = "    !!!    " + xlWorkSheet2.Cells[i, 1].Text + " "
                        + xlWorkSheet2.Cells[i, 2].Text + "     !!!     ";
                    xlWorkSheet2.Cells[i, 8] = "0";
                    xlWorkSheet2.Cells[i, 9] = "0";
                    plati = false;
                }
                if (xlWorkSheet2.Cells[i + 1, 1].Text == "")
                    break; // Zabrání aplikaci konrolovat nevyplněné položky
            }

            if (plati) return true; else return false;
        }

        /// <summary>
        /// Metoda, která odstraní diakritiku při tvorbě názvu fotografií
        /// </summary>
        /// <param name="text"></param>
        /// <returns>Původní string bez diakritiky</returns>
        static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }

        /// <summary>
        /// Výběr složky
        /// </summary>
        /// <returns>Cestu ke složce</returns>
        private string vyberSlozku()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowDialog();
            string result = fbd.SelectedPath;

            return result;
        }

        /// <summary>
        /// Výběr Word Template
        /// </summary>
        /// <returns>Adresu vybraného souboru</returns>
        private string vyberSoubor()
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                string result = openFileDialog2.FileName;
                opened = true;
                return result;
            }
            else
            {
                opened = false;
                return null;
            }
        }

        /// <summary>
        /// Vyčištění paměti od excel souborů
        /// </summary>
        /// <param name="obj">Aplikace</param>
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// Vypnutí aplikace přes MenuStrip
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ukončitProgramToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OnApplicationExit(this, null);
            Application.Exit();
        }

        /// <summary>
        /// Event zavolaný při ukončení aplikace, vypíná běžící Excel a Word aplikace
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnApplicationExit(object sender, EventArgs e)
        {
            try
            {
                xlWorkBook2.Close(false, null, null);
            }
            catch
            {

            }

            try
            {
                xlWorkBook.Close(false, null, null);
            }
            catch
            {

            }

            try
            {
                xlApp.Quit();
            }
            catch
            {

            }

            try
            {
                xlApp2.Quit();
            }
            catch
            {

            }

            try
            {
                aDoc.Close(false, ref missing, ref missing);
            }
            catch
            {

            }

            try
            {
                wordApp.Quit();
            }
            catch
            {

            }
        }

        /// <summary>
        /// Metoda, která vypíše kód chyby na obrazovku
        /// </summary>
        /// <param name="Ex"></param>
        private void exceptionsHandler(Exception Ex)
        {
            MessageBox.Show(Ex.Message, "Chyba", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Error);
        }
        //Konec Part 4

        //Nevyužité
        private void Průkazkovač_Load(object sender, EventArgs e)
        {

        }
        private void souborToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void label7_Click(object sender, EventArgs e)
        {

        }
    }
}
