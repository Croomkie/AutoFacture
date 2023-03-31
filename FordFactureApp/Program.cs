using System.Globalization;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace FordFactureApp // Note: actual namespace depends on the project name.
{
    internal class Program
    {

        private static string? datefacture = "";
        private static string? numFacture = "";

        private static string? nomClient = "";
        private static string? adresseClient = "";
        private static string? telClient = "";
        private static string? kilometrageVoiture = "";
        private static string? dateTravaux = "";
        static void Main(string[] args)
        {
            string templatePath = "C:\\Users\\ldp03\\Documents\\EntretienFiesta\\WordModele\\fordFactureModele.docx";
            string outputPath = "C:\\Users\\ldp03\\Documents\\EntretienFiesta\\WordModele\\fordFactureModeleModified.docx";

            AskValueOfVarible();

            EditWordTemplate(templatePath, outputPath);
        }

        private static void AskValueOfVarible()
        {
            Console.WriteLine("Date de la facture ?");
            datefacture = Console.ReadLine();
            Console.WriteLine("Numéro de la facture ?");
            numFacture = Console.ReadLine();

            Console.WriteLine("Quelle est le nom du client ?");
            nomClient = Console.ReadLine();
            Console.WriteLine("Adresse du client ?");
            adresseClient = Console.ReadLine();
            Console.WriteLine("Tel du client ?");
            telClient = Console.ReadLine();
            Console.WriteLine("Kilometrage de la voiture ?");
            kilometrageVoiture = Console.ReadLine();
            Console.WriteLine("Date réalisation travaux ?");
            dateTravaux = Console.ReadLine();
        }

        private static void EditWordTemplate(string templatePath, string outputPath)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = null;
            object missing = Type.Missing;

            try
            {
                // Ouvrir le modèle Word
                wordDoc = wordApp.Documents.Open(templatePath);

                // Date num facture
                FindAndReplace(wordDoc, "#datefacture", DateTime.Now.ToString("D", new CultureInfo("fr-FR")));
                FindAndReplace(wordDoc, "#NUMFACTURE", numFacture);


                // Remplacer les variables client
                FindAndReplace(wordDoc, "#nomClient", nomClient);
                FindAndReplace(wordDoc, "#adresseClient", adresseClient);
                FindAndReplace(wordDoc, "#telClient", telClient);
                FindAndReplace(wordDoc, "#kilometrageVoiture", kilometrageVoiture);
                FindAndReplace(wordDoc, "#dateTravaux", DateTime.Now.ToString("dd MMMM, yyyy", new CultureInfo("fr-FR")));
  
                // Trouver le tableau dans le document
                Word.Table table = null;
                foreach (Word.Table tbl in wordDoc.Tables)
                {
                    if (tbl.Title == "TableauProduit")
                    {
                        table = tbl;
                        break;
                    }
                }

                // Ajouter et remplir des lignes
                if (table != null)
                {
                    string styleName = "TableauFacture";
                    object styleNameObj = styleName;
                    Word.Row firstRow = table.Rows[1];
                    foreach (Word.Cell cell in firstRow.Cells)
                    {
                        cell.Range.set_Style(ref styleNameObj);
                    }

                    for (int i = 0; i < 5; i++)
                    {
                        Word.Row newRow = table.Rows.Add(ref missing);

                        newRow.Cells[1].Range.Text = "qteProduit";
                        newRow.Cells[2].Range.Text = "redArticle";
                        newRow.Cells[3].Range.Text = "descriptionProduit";
                        newRow.Cells[4].Range.Text = "prixUnitaire";
                        newRow.Cells[5].Range.Text = "totalLigne";
                    }

                    Word.Row lastRow = table.Rows.Add(ref missing);
                    lastRow.Cells[1].Range.Text = null;
                    lastRow.Cells[2].Range.Text = null;
                    lastRow.Cells[3].Range.Text = null;
                    lastRow.Cells[4].Range.Text = null;
                    lastRow.Cells[5].Range.Text = "187";
                }

                // Sauvegarder le document
                // Sauvegarder le document modifié sous un nouveau nom
                object outputFileName = outputPath;
                wordDoc.SaveAs2(ref outputFileName);

                Console.WriteLine("Le document modifié a été enregistré.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erreur : " + ex.Message);
            }
            finally
            {
                // Fermer le document Word et quitter Word
                if (wordDoc != null)
                {
                    wordDoc.Close(ref missing, ref missing, ref missing);
                    Marshal.ReleaseComObject(wordDoc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit(ref missing, ref missing, ref missing);
                    Marshal.ReleaseComObject(wordApp);
                }
            }
        }

        static void FindAndReplace(Word.Document document, string findTextString, string replaceText)
        {
            object findText = findTextString;
            object replaceWith = replaceText;
            object replace = Word.WdReplace.wdReplaceOne;
            object missing = Type.Missing;

            Word.Range docRange = document.Content;

            docRange.Find.ClearFormatting();
            docRange.Find.Execute(ref findText, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith, ref replace, ref missing, ref missing, ref missing, ref missing);
        }


    }
}