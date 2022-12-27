using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForTest3
{
    public static class Table
    {
        static string path = @"C:\Users\Михаил\Desktop\test6.docx";
        public static void TabCreate()
        {


            using (WordprocessingDocument wordDoc2 = WordprocessingDocument.Open(path, true))
            {
               

                MainDocumentPart mainPart = wordDoc2.MainDocumentPart;
                DocumentSettingsPart settingsPart = mainPart.DocumentSettingsPart;
                DocumentVariables docVars = settingsPart.Settings.Descendants<DocumentVariables>().FirstOrDefault();
                if (docVars == null)
                {
                    settingsPart.Settings = new Settings(new DocumentVariables());
                }
                DocumentVariable docVar = new DocumentVariable() { Name = "Var13", Val = "TestVar13" };




                var doc = wordDoc2.MainDocumentPart.Document;
                DocumentFormat.OpenXml.Wordprocessing.Table table = doc.MainDocumentPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault();
               
                int icounterfortableservice;
                for (icounterfortableservice = 0; icounterfortableservice < 1; icounterfortableservice++)
                {
                    DocumentFormat.OpenXml.Wordprocessing.TableRow tr = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                    DocumentFormat.OpenXml.Wordprocessing.TableCell tablecellService1 = new DocumentFormat.OpenXml.Wordprocessing.TableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Василий"))));
                    DocumentFormat.OpenXml.Wordprocessing.TableCell tablecellService2 = new DocumentFormat.OpenXml.Wordprocessing.TableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text("Петрович"))));
                    DocumentFormat.OpenXml.Wordprocessing.TableCell tablecellService3 = new DocumentFormat.OpenXml.Wordprocessing.TableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(docVar.Val))));

                     tr.Append(tablecellService1, tablecellService2, tablecellService3);
                    table.AppendChild(tr);

                }
                settingsPart.Settings.Save();
                wordDoc2.Save();

            }
            using (WordprocessingDocument document = WordprocessingDocument.Open(path, true))
            {
                DocumentSettingsPart settingsPart = document.MainDocumentPart.GetPartsOfType<DocumentSettingsPart>().First();
                // Create object to update fields on open
                UpdateFieldsOnOpen updateFields = new UpdateFieldsOnOpen();
                updateFields.Val = new DocumentFormat.OpenXml.OnOffValue(true);
                // Insert object into settings part.
                settingsPart.Settings.PrependChild<UpdateFieldsOnOpen>(updateFields);
                settingsPart.Settings.Save();
            }
        }



        //возвращает значение docVarable
        public static DocumentVariable GetVariableByName(string name, WordprocessingDocument document)
        {
            //var o = g.InnerXml;
            // Console.WriteLine(o);
            // Get the document settings part
            DocumentSettingsPart documentSettings = document.MainDocumentPart.DocumentSettingsPart;
            // Get the settings element
            Settings settings = documentSettings.Settings;
            // Get the DocumentVariables element
            DocumentVariables variables = settings.Elements<DocumentVariables>().FirstOrDefault();
            // check if the variables are not null
            if (variables != null)
            {
                return variables.Elements<DocumentVariable>().Where(v => v.Name == name)
                    .FirstOrDefault();
            }
            return null;
        }
    }
}
