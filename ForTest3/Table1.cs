using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ForTest3
{
    public static class Table1
    {
        static string path = @"C:\Users\Михаил\Desktop\test6.docx";

        public static void Table()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(path, true))
            {
                DocumentFormat.OpenXml.Wordprocessing.Table myTable = doc.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().First();
                //myTable.LocalName = "";
                TableRow theRow = myTable.Elements<TableRow>().Last();
                for (int i = 0; i < 1; i++)
                {
                    TableRow rowCopy = (TableRow)theRow.CloneNode(true);

                  

                    TableCell cell = rowCopy.Elements<TableCell>().ElementAt(0);
                    TableCell cell1 = rowCopy.Elements<TableCell>().ElementAt(1);


                    Paragraph p = cell.Elements<Paragraph>().First();
                    Run r = p.Elements<Run>().First();
                    Text t = r.Elements<Text>().First();
                    t.Text = "Изменена 1 ячейка";

                    Paragraph p1 = cell1.Elements<Paragraph>().First();
                    Run r1 = p1.Elements<Run>().First();
                    Text t1 = r1.Elements<Text>().First();
                    t1.Text = "Изменена 2 ячейка";

                    myTable.AppendChild(rowCopy);
                }
                 //myTable.RemoveChild(theRow);
            }
        }

    }
}
