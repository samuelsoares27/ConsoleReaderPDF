

using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using IronXL;


var path = "C:\\Trabalho\\Arquivos\\prefeitura\\alvarahabilite";
List<string> list = new ();
//bool habiteses = false;
//bool alvara = false;
var contador = 1;

string[] arquivos = Directory.GetFiles(path);

#region Pegando arquivos e salvando em lista
foreach (string arq in arquivos)
{
    using PdfReader leitor = new(arq);
    for (int i = 1; i <= leitor.NumberOfPages; i++)
    {
        if(i == 1)
            list.Add(PdfTextExtractor.GetTextFromPage(leitor, i));
    }

}

foreach (var item in list)
{
    if (item.Contains("HABITE-SE"))
    {
        System.Console.WriteLine("linha: " + contador + " HABITE-SE");
    }
    if (item.Contains("ALVARÁ"))
    {
        System.Console.WriteLine("linha: " + contador + " ALVARÁ");
    }
    contador++;
}
#endregion

#region habitese

WorkBook workbookHabitese = WorkBook.Create(ExcelFileFormat.XLSX);
var sheetHabitese = workbookHabitese.CreateWorkSheet("Result Sheet");
sheetHabitese["A1"].Value = "TITULO";
sheetHabitese["B1"].Value = "ENDEREÇO DA OBRA";
sheetHabitese["C1"].Value = "BAIRRO";
sheetHabitese["D1"].Value = "CIDADE";
sheetHabitese["E1"].Value = "PROPRIETÁRIO DO IMÓVEL (NOME)";
sheetHabitese["F1"].Value = "PROPRIETÁRIO DO IMÓVEL (CPF)";
sheetHabitese["G1"].Value = "RESPONSÁVEL PELA EXECUÇÃO DA OBRA (NOME)";
sheetHabitese["H1"].Value = "RESPONSÁVEL PELA EXECUÇÃO DA OBRA (CPF)";
sheetHabitese["I1"].Value = "RESPONSÁVEL TÉCNICO";
sheetHabitese["J1"].Value = "DESCRICAO TERRENO";
sheetHabitese["K1"].Value = "ESPECIFICAÇÃO (TIPO DE HABITE-SE)";
sheetHabitese["L1"].Value = "ESPECIFICAÇÃO (DETALHES)";
sheetHabitese["M1"].Value = "AREAS PRINCIPAIS";
sheetHabitese["N1"].Value = "AREA TOTAL";
sheetHabitese["O1"].Value = "OBSERVACAO";
sheetHabitese["P1"].Value = "CIDADE/DATA";


workbookHabitese.SaveAs($"{path}\\ExcelHabitese.xlsx");

#endregion

#region alvara

WorkBook workbookAlvara = WorkBook.Create(ExcelFileFormat.XLSX);
var sheetAlvara = workbookAlvara.CreateWorkSheet("Result Sheet");
sheetAlvara["A1"].Value = "TITULO";
sheetAlvara["B1"].Value = "PROPRIETÁRIO DO IMÓVEL (NOME)";
sheetAlvara["C1"].Value = "PROPRIETÁRIO DO IMÓVEL (CPF)";
sheetAlvara["D1"].Value = "AUTOR PROJETO (NOME)";
sheetAlvara["E1"].Value = "AUTOR PROJETO (CPF)";
sheetAlvara["F1"].Value = "AUTOR PROJETO (CREA)";
sheetAlvara["G1"].Value = "RESPONSÁVEL TÉCNICO (NOME)";
sheetAlvara["H1"].Value = "RESPONSÁVEL TÉCNICO (CPF)";
sheetAlvara["I1"].Value = "RESPONSÁVEL TÉCNICO (CREA)";
sheetAlvara["J1"].Value = "CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA (NOME)";
sheetAlvara["K1"].Value = "CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA (CPF/CNPJ)";
sheetAlvara["L1"].Value = "DESCRICAO";
sheetAlvara["M1"].Value = "AREAS PRINCIPAIS";
sheetAlvara["N1"].Value = "ESPECIFICAÇÃO (DETALHES)";
sheetAlvara["O1"].Value = "OBSERVACAO";
sheetAlvara["P1"].Value = "LEI";

workbookAlvara.SaveAs($"{path}\\ExcelAlvara.xlsx");

#endregion


#region salvando código
/*
if (list[i - 1].Contains("HABITE-SE"))
{
    habiteses = true;
}
if (list[i - 1].Contains("ALVARÁ"))
{
    alvara = true;
}
if (habiteses)
{
    System.Console.WriteLine("HABITE-SE");
}

if (alvara)
{
    System.Console.WriteLine("ALVARÁ");
}

for (var index = 0; index < list.Count; index++)
{
    System.Console.WriteLine("linha: " + index + " texto: " + list[index]);
}
*/
#endregion