using Microsoft.Office.Interop.Word;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Excel = Microsoft.Office.Interop.Excel;

namespace CreateExcelEscrime
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBoxDestinataire.Items.Clear();
            comboBoxDestinataire.Items.Add("A.R.A.B.");
            comboBoxDestinataire.Items.Add("Assurance DAS");
            comboBoxDestinataire.Items.Add("Baptiste Motte");
            comboBoxDestinataire.Items.Add("Club");
            comboBoxDestinataire.Items.Add("Ecole Don Bosco");
            comboBoxDestinataire.Items.Add("Eliot Brabant");
            comboBoxDestinataire.Items.Add("Etat du compte");
            comboBoxDestinataire.Items.Add("Fabrice Razanajao");
            comboBoxDestinataire.Items.Add("Félix Trannoy");
            comboBoxDestinataire.Items.Add("Frais bancaires");
            comboBoxDestinataire.Items.Add("Ligue francophone");
            comboBoxDestinataire.Items.Add("Oscar Deblocq");
            comboBoxDestinataire.Items.Add("Planète Escrime");

            comboBoxPériode.Items.Add("2025-2026");
            comboBoxPériode.SelectedIndex = 0;

            comboBoxRaison.Items.Add("Achat Matériel pour le club");
            comboBoxRaison.Items.Add("Achat Matériel pour un tireur");
            comboBoxRaison.Items.Add("Location Matériel");
            comboBoxRaison.Items.Add("Payement à l'année");
            comboBoxRaison.Items.Add("Payement 1er Carte");
            comboBoxRaison.Items.Add("Payement 2ème Carte");
            comboBoxRaison.Items.Add("Payement 3ème Carte");
            comboBoxRaison.Items.Add("Payement 4ème Carte");
            comboBoxRaison.Items.Add("Payement 5ème Carte");
            comboBoxRaison.Items.Add("Payement 6ème Carte");
            comboBoxRaison.Items.Add("Payement de la salle");
            comboBoxRaison.Items.Add("Remboursement location matériel à un tireur");

            label5.Text = DateTime.Now.ToShortDateString();

            RemplissagePourASBL();
        }

        private void RemplissagePourASBL()
        {
            comboBoxDépensesRecettes.Items.Clear();
            comboBoxPots.Items.Clear();

            if (textBoxMontant.Text == "" || textBoxMontant.Text == "-")
            {

            }
            else
            {
                if (Convert.ToInt32(textBoxMontant.Text) > 0)
                {
                    comboBoxDépensesRecettes.Items.Add("Recettes");
                    comboBoxDépensesRecettes.SelectedIndex = 0;

                    comboBoxPots.Items.Add("");
                    comboBoxPots.Items.Add("Autres recettes");
                    comboBoxPots.Items.Add("Cotisations");
                    comboBoxPots.Items.Add("LocationMatériel");
                    comboBoxPots.Items.Add("Subsides");
                    comboBoxPots.Items.Add("VenteMatériel");
                }
                else
                {
                    comboBoxDépensesRecettes.Items.Add("Dépenses");
                    comboBoxDépensesRecettes.SelectedIndex = 0;

                    comboBoxPots.Items.Add("");
                    comboBoxPots.Items.Add("asbl");
                    comboBoxPots.Items.Add("Assurance");
                    comboBoxPots.Items.Add("Autres dépenses");
                    comboBoxPots.Items.Add("Ligue");
                    comboBoxPots.Items.Add("Marchandises");
                    comboBoxPots.Items.Add("Rémunérations");
                    comboBoxPots.Items.Add("Site");
                    comboBoxPots.Items.Add("Salle");
                }
            }
        }

        private void btnComptabiliser_Click(object sender, EventArgs e)
        {
            if (comboBoxPots.Text == "")
            {
                return;
            }

            btnComptabiliser.Enabled = !btnComptabiliser.Enabled;
            string path = Path.GetDirectoryName(Application.ExecutablePath) + @"\CreateExcel.txt";

            string rowToAdd = monthCalendar1.SelectionRange.Start.ToString("yyyy-MM-dd") + ";"
                + comboBoxDestinataire.Text + ";" + textBoxMontant.Text + ";" + comboBoxRaison.Text + ";" + comboBoxPériode.Text;

            InsertText(path, rowToAdd);
            btnComptabiliser.Enabled = !btnComptabiliser.Enabled;
        }

        public static void InsertText(string path, string newText)
        {
            if (File.Exists(path))
            {
                string oldText = File.ReadAllText(path);
                using (var sw = new StreamWriter(path, false))
                {
                    sw.WriteLine(newText);
                    sw.WriteLine(oldText);
                }
            }
            else File.WriteAllText(path, newText);

            string myFileData = File.ReadAllText(path);
            File.WriteAllText(path, myFileData.TrimEnd(Environment.NewLine.ToCharArray()));
        }

        private void btnCreateExcel_Click(object sender, EventArgs e)
        {
            btnCreateExcel.Enabled = !btnCreateExcel.Enabled;

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Date";
            xlWorkSheet.Cells[1, 2] = "Destinataire";
            xlWorkSheet.Cells[1, 3] = "Montant Positif";
            xlWorkSheet.Cells[1, 4] = "Montant Négatif";
            xlWorkSheet.Cells[1, 5] = "Motif";
            xlWorkSheet.Cells[1, 6] = "Période";
            xlWorkSheet.Cells[1, 7] = "Etat du compte";
            xlWorkSheet.Cells[1, 8] = "Dépenses/Recettes";
            xlWorkSheet.Cells[1, 9] = "Pots";

            string path = Path.GetDirectoryName(Application.ExecutablePath) + @"\CreateExcel.txt";
            string[] rows = File.ReadAllLines(path);

            int rowIndew = 3;
            foreach (string row in rows)
            {
                string[] cellule = row.Split(';');

                xlWorkSheet.Cells[rowIndew, 1] = cellule[0];
                xlWorkSheet.Cells[rowIndew, 1].NumberFormat = "[$-fr-BE]jjjj j mmmm aaaa";
                xlWorkSheet.Cells[rowIndew, 2] = cellule[1];

                if (cellule[1] == "Etat du compte")
                {
                    xlWorkSheet.Cells[rowIndew, 7] = cellule[2];
                }
                else
                {
                    if (decimal.Parse(cellule[2]) > 0)
                    {
                        xlWorkSheet.Cells[rowIndew, 3] = cellule[2];
                    }
                    else
                    {
                        if (decimal.Parse(cellule[2]) < 0)
                        {

                            xlWorkSheet.Cells[rowIndew, 4] = cellule[2];
                        }

                    }
                    xlWorkSheet.Cells[rowIndew, 5] = cellule[3];
                    xlWorkSheet.Cells[rowIndew, 6] = cellule[4];

                    xlWorkSheet.Cells[rowIndew, 8] = cellule[5];
                    xlWorkSheet.Cells[rowIndew, 9] = cellule[6];
                }

                rowIndew++;
            }


            Excel.Range usedrange = xlWorkSheet.UsedRange;
            usedrange.Columns.AutoFit();

            usedrange.Style.HorizontalAlignment = -4108;

            string pathOutput = Path.GetDirectoryName(Application.ExecutablePath) + @"\ExportCreateExcel.xlsx";
            File.Delete(pathOutput);

            xlWorkBook.SaveAs(pathOutput, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            System.Diagnostics.Process.Start(pathOutput);
            btnCreateExcel.Enabled = !btnCreateExcel.Enabled;
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            label5.Text = monthCalendar1.SelectionStart.Date.ToLongDateString() + "  (" + monthCalendar1.SelectionStart.Date.ToShortDateString() + ")";
        }

        private void btnCreateWord_Click(object sender, EventArgs e)
        {
            string path = Path.GetDirectoryName(Application.ExecutablePath) + @"\CreateExcel.txt";
            string[] rows = File.ReadAllLines(path);

            btnCreateWord.Enabled = !btnCreateWord.Enabled;

            Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
            winword.ShowAnimation = false;
            winword.Visible = false;
            object missing = System.Reflection.Missing.Value;
            Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            Paragraph paragraph1 = document.Content.Paragraphs.Add(ref missing);
            object styleHeading1 = "Heading 1";
            //paragraph1.Range.set_Style(ref styleHeading1);

            paragraph1.Range.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

            Table firstTable = document.Tables.Add(paragraph1.Range, rows.Count(), 7, ref missing, ref missing);
            firstTable.Borders.Enable = 1;



            firstTable.Cell(1, 1).Range.Text = "Date";
            firstTable.Cell(1, 2).Range.Text = "Destinataire";
            firstTable.Cell(1, 3).Range.Text = "Montant Positif";
            firstTable.Cell(1, 4).Range.Text = "Montant Négatif";
            firstTable.Cell(1, 5).Range.Text = "Motif";
            firstTable.Cell(1, 6).Range.Text = "Période";
            firstTable.Cell(1, 7).Range.Text = "Etat du compte";

            int rowIndew = 3;
            foreach (string row in rows)
            {
                string[] cellule = row.Split(';');
                var n = DateTime.Parse(cellule[0]).ToString("dddd dd-MMMM-yy");

                firstTable.Cell(rowIndew, 1).Range.Text = FormatDay(n);
                firstTable.Cell(rowIndew, 2).Range.Text = cellule[1];

                if (cellule[1] == "Etat du compte")
                {
                    firstTable.Cell(rowIndew, 7).Range.Text = cellule[2];
                }
                else
                {
                    if (decimal.Parse(cellule[2]) > 0)
                    {
                        firstTable.Cell(rowIndew, 3).Range.Text = cellule[2];
                    }
                    else
                    {
                        if (decimal.Parse(cellule[2]) < 0)
                        {

                            firstTable.Cell(rowIndew, 4).Range.Text = cellule[2];
                        }

                    }
                }

                firstTable.Cell(rowIndew, 5).Range.Text = cellule[3];
                firstTable.Cell(rowIndew, 6).Range.Text = cellule[4];

                firstTable.Cell(rowIndew, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                firstTable.Cell(rowIndew, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                firstTable.Cell(rowIndew, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                firstTable.Cell(rowIndew, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                firstTable.Cell(rowIndew, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                firstTable.Cell(rowIndew, 6).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                firstTable.Cell(rowIndew, 1).Range.Font.Size = 10;

                rowIndew++;
            }

            firstTable.Range.Columns[7].AutoFit();
            firstTable.Range.Columns[6].AutoFit();
            firstTable.Range.Columns[5].AutoFit();
            firstTable.Range.Columns[4].AutoFit();
            firstTable.Range.Columns[3].AutoFit();
            firstTable.Range.Columns[2].AutoFit();
            firstTable.Range.Columns[1].AutoFit();




            object filename = Path.GetDirectoryName(Application.ExecutablePath) + @"\ExportCreateWord.docx";
            document.SaveAs2(ref filename);
            document.Close(ref missing, ref missing, ref missing);
            document = null;
            winword.Quit(ref missing, ref missing, ref missing);
            winword = null;
            // MessageBox.Show("Document created successfully !");
            System.Diagnostics.Process.Start(filename.ToString());

            btnCreateWord.Enabled = !btnCreateWord.Enabled;
        }

        private string FormatDay(string n)
        {
            n = n.Replace("lundi", "Lun ");
            n = n.Replace("mardi", "Mar ");
            n = n.Replace("mercredi", "Mer ");
            n = n.Replace("jeudi", "Jeu ");
            n = n.Replace("vendredi", "Ven ");
            n = n.Replace("samedi", "Sam ");
            n = n.Replace("dimanche", "Dim ");
            return n;
        }

        private void textBoxMontant_TextChanged(object sender, EventArgs e)
        {
            RemplissagePourASBL();
        }
    }
}
