// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using static System.Net.WebRequestMethods;

Console.WriteLine("Hello, World!");
string path = @"C:\Users\Михаил\Desktop\служебка на создание сайта.docx";

//UbdateDocVar("Var4", "Михаил Андреевич");
//UbdateDocVar("Var12", "вАСЯ Михаил Андреевич");
//ForTest3.Table1.Table();
//ForTest3.Table.TabCreate();
using (WordprocessingDocument pkg = WordprocessingDocument.Open(path, true))
{
  MainDocumentPart mainPart = pkg.MainDocumentPart;
        DocumentSettingsPart settingsPart = mainPart.DocumentSettingsPart;
        DocumentVariables docVars = settingsPart.Settings.Descendants<DocumentVariables>().FirstOrDefault();
        if (docVars == null)
        {
            settingsPart.Settings = new Settings(new DocumentVariables());
        }
    /* DocumentVariable docVar = new DocumentVariable() { Name = "dd", Val = "dd" };
     DocumentVariable docVar1 = new DocumentVariable() { Name = "mm", Val = "mm" };
     DocumentVariable docVar2 = new DocumentVariable() { Name = "yy", Val = "yy" };
     docVars.Append(docVar);
     docVars.Append(docVar1);
     docVars.Append(docVar2);*/

    /* DocumentVariable docVar3 = new DocumentVariable() { Name = "NameOrganization", Val = "NameOrganization" };
     docVars.Append(docVar3);
     DocumentVariable docVar4 = new DocumentVariable() { Name = "AddresseeFio", Val = "AddresseeFio" };
     docVars.Append(docVar4);
     DocumentVariable docVar5 = new DocumentVariable() { Name = "AddresseePost", Val = "AddresseePost" };
     docVars.Append(docVar5);
     DocumentVariable docVar6 = new DocumentVariable() { Name = "ResponsibleFio", Val = "ResponsibleFio" };
     docVars.Append(docVar6);
     DocumentVariable docVar7 = new DocumentVariable() { Name = "ResponsiblePost", Val = "ResponsiblePost" };
     docVars.Append(docVar7);
    */
    DocumentVariable docVar4 = new DocumentVariable() { Name = "AddresseeFio", Val = "AddresseeFio" };
    docVars.Append(docVar4);
    /* DocumentVariable variable = GetVariableByName("Var1", pkg);
     if (variable == null)
     {

         DocumentSettingsPart settingsPart = pkg.MainDocumentPart.DocumentSettingsPart;
         settingsPart.Settings = new Settings(new DocumentVariables(
             new DocumentVariable() { Name = "Var1", Val = "Test" })) ;
         Console.WriteLine("Создание переменной");
         settingsPart.Settings.Save();
         pkg.Save();
     }*/

}



UbdateDocVar("IP", "вАСЯ Михаил Андреевич");
UbdateDocVar("FI", "вАСЯ М.А.");

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




//обновляет значение docVarable
void UbdateDocVar(string var, string val)
{
    using (WordprocessingDocument document = WordprocessingDocument.Open(path, true))
    {
        DocumentVariable variable = GetVariableByName(var, document);
        if (variable != null)
            Console.WriteLine($"перед изменением переменной: {variable.Val}");
        SetDocumentVariableValue(variable, val);

        document.Save();
        Console.WriteLine($"после изменением переменной: {variable.Val}");
    }
}

//возвращает значение docVarable
DocumentVariable GetVariableByName(string name, WordprocessingDocument document)
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

//меняет значение docVarable
void SetDocumentVariableValue(DocumentVariable variable, string value)
{
    variable.Val = value;
}














