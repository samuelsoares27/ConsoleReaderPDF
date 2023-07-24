

using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using IronXL;
using ConsoleReaderPDF;

var path = "C:\\Trabalho\\Arquivos\\prefeitura\\alvarahabilite";
List<Objeto> list = new ();
string[] arquivos = Directory.GetFiles(path);

#region Pegando arquivos e salvando em lista

foreach (string arq in arquivos)
{
    using PdfReader leitor = new(arq);
    for (int i = 1; i <= leitor.NumberOfPages; i++)
    {
        if(i == 1)
        {
            Objeto objeto = new();
            objeto.Nome = System.IO.Path.GetFileName(arq);
            objeto.Texto = PdfTextExtractor.GetTextFromPage(leitor, i).ToString();
            list.Add(objeto);
        }
            
    }

}

#endregion

#region habitese

WorkBook workbookHabitese = WorkBook.Create(ExcelFileFormat.XLSX);
var sheetHabitese = workbookHabitese.CreateWorkSheet("HABITE-SE");

sheetHabitese["A1"].Value = "NOME ARQUIVO";
sheetHabitese["B1"].Value = "TITULO";
sheetHabitese["C1"].Value = "ENDEREÇO DA OBRA";
sheetHabitese["D1"].Value = "BAIRRO";
sheetHabitese["E1"].Value = "CIDADE";
sheetHabitese["F1"].Value = "PROPRIETÁRIO DO IMÓVEL (NOME)";
sheetHabitese["G1"].Value = "PROPRIETÁRIO DO IMÓVEL (CPF/CNPJ)";
sheetHabitese["H1"].Value = "RESPONSÁVEL PELA EXECUÇÃO DA OBRA (NOME)";
sheetHabitese["I1"].Value = "RESPONSÁVEL PELA EXECUÇÃO DA OBRA (CPF/CNPJ)";
sheetHabitese["J1"].Value = "RESPONSÁVEL TÉCNICO";
sheetHabitese["K1"].Value = "DESCRICAO TERRENO";
sheetHabitese["L1"].Value = "ESPECIFICAÇÃO (TIPO DE HABITE-SE)";
//sheetHabitese["M1"].Value = "ESPECIFICAÇÃO (DETALHES)";
//sheetHabitese["N1"].Value = "AREAS PRINCIPAIS";
sheetHabitese["O1"].Value = "AREA TOTAL";
sheetHabitese["P1"].Value = "OBSERVACAO";


var linha = 2;
var listaHabitese = list.Where(tipo => tipo.Texto.Contains("ENDEREÇO DA OBRA:")).ToList();
foreach (var row in listaHabitese)
{
    sheetHabitese[$"A{linha}"].Value = $"{row.Nome}";    
    sheetHabitese[$"B{linha}"].Value = $"HABITE-SE {row.Texto.Substring(row.Texto.IndexOf("Nº"), row.Texto.IndexOf("ENDEREÇO") - row.Texto.IndexOf("Nº"))}".Replace("\n", " ");
    sheetHabitese[$"C{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("ENDEREÇO DA OBRA:"), row.Texto.IndexOf("BAIRRO:") - row.Texto.IndexOf("ENDEREÇO DA OBRA:"))}".Replace("ENDEREÇO DA OBRA:", "").Replace("\n", " ");
    sheetHabitese[$"D{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("BAIRRO:"), row.Texto.IndexOf("CIDADE:") - row.Texto.IndexOf("BAIRRO:"))}".Replace("BAIRRO:", "").Replace("\n", " ");
    sheetHabitese[$"E{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("CIDADE:"), row.Texto.IndexOf("PROPRIETÁRIO DO IMÓVEL:") - row.Texto.IndexOf("CIDADE:"))}".Replace("CIDADE:", "").Replace("\n", " ");

    var ProprietarioImovel = $"{row.Texto.Substring(row.Texto.IndexOf("PROPRIETÁRIO DO IMÓVEL:"), row.Texto.IndexOf("RESPONSÁVEL PELA EXECUÇÃO DA OBRA:") - row.Texto.IndexOf("PROPRIETÁRIO DO IMÓVEL:"))}";
    if(ProprietarioImovel.IndexOf("CNPJ:") != -1)
    {
        sheetHabitese[$"F{linha}"].Value = $"{ProprietarioImovel.Substring(ProprietarioImovel.IndexOf("NOME:"), ProprietarioImovel.IndexOf("CNPJ:") - ProprietarioImovel.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetHabitese[$"G{linha}"].Value = $"{ProprietarioImovel.Substring(ProprietarioImovel.IndexOf("CNPJ:"), ProprietarioImovel.Length - ProprietarioImovel.IndexOf("CNPJ:"))}".Replace("CNPJ:", "").Replace("\n", " ");
    }
    else if (ProprietarioImovel.IndexOf("CPF:") != -1)
    {
        sheetHabitese[$"F{linha}"].Value = $"{ProprietarioImovel.Substring(ProprietarioImovel.IndexOf("NOME:"), ProprietarioImovel.IndexOf("CPF:") - ProprietarioImovel.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetHabitese[$"G{linha}"].Value = $"{ProprietarioImovel.Substring(ProprietarioImovel.IndexOf("CPF:"), ProprietarioImovel.Length - ProprietarioImovel.IndexOf("CPF:"))}".Replace("CPF:", "").Replace("\n", " ");

    }

    var ResponsavelExecucaoObra = $"{row.Texto.Substring(row.Texto.IndexOf("RESPONSÁVEL PELA EXECUÇÃO DA OBRA:"), row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:") - row.Texto.IndexOf("RESPONSÁVEL PELA EXECUÇÃO DA OBRA:"))}";
    if (ResponsavelExecucaoObra.IndexOf("CNPJ:") != -1)
    {
        sheetHabitese[$"H{linha}"].Value = $"{ResponsavelExecucaoObra.Substring(ResponsavelExecucaoObra.IndexOf("NOME:"), ResponsavelExecucaoObra.IndexOf("CNPJ:") - ResponsavelExecucaoObra.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetHabitese[$"I{linha}"].Value = $"{ResponsavelExecucaoObra.Substring(ResponsavelExecucaoObra.IndexOf("CNPJ:"), ResponsavelExecucaoObra.Length - ResponsavelExecucaoObra.IndexOf("CNPJ:"))}".Replace("CNPJ:", "").Replace("\n", " ");

    }else if (ResponsavelExecucaoObra.IndexOf("CPF:") != -1)
    {
        sheetHabitese[$"H{linha}"].Value = $"{ResponsavelExecucaoObra.Substring(ResponsavelExecucaoObra.IndexOf("NOME:"), ResponsavelExecucaoObra.IndexOf("CPF:") - ResponsavelExecucaoObra.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetHabitese[$"I{linha}"].Value = $"{ResponsavelExecucaoObra.Substring(ResponsavelExecucaoObra.IndexOf("CPF:"), ResponsavelExecucaoObra.Length - ResponsavelExecucaoObra.IndexOf("CPF:"))}".Replace("CPF:", "").Replace("\n", " ");

    }


    sheetHabitese[$"J{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:"), row.Texto.IndexOf("Conforme") - row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:"))}".Replace("RESPONSÁVEL TÉCNICO:", "").Replace("\n", " ");
    sheetHabitese[$"K{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("Conforme"), row.Texto.IndexOf("ESPECIFICAÇÃO:") - row.Texto.IndexOf("Conforme"))}".Replace("\n", " ");
    
    var especificacoes = $"{row.Texto.Substring(row.Texto.IndexOf("ESPECIFICAÇÃO:"), row.Texto.IndexOf("Área total da obra:") - row.Texto.IndexOf("ESPECIFICAÇÃO:"))}".Replace("\n", " ");
    sheetHabitese[$"L{linha}"].Value = $"{especificacoes.Substring(especificacoes.IndexOf("TIPO DE HABITE-SE:"), especificacoes.IndexOf("Dados da obra:") - especificacoes.IndexOf("TIPO DE HABITE-SE:"))}".Replace("TIPO DE HABITE-SE:", "").Replace("\n", " ");
    //sheetHabitese[$"M{linha}"].Value = "";
    //sheetHabitese[$"N{linha}"].Value = "";
    sheetHabitese[$"O{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("Área total da obra:"), row.Texto.IndexOf("OBSERVAÇÃO:") - row.Texto.IndexOf("Área total da obra:"))}".Replace("Área total da obra:", "").Replace("\n", " ");
    sheetHabitese[$"P{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("OBSERVAÇÃO:"), row.Texto.IndexOf("MG, em") - row.Texto.IndexOf("OBSERVAÇÃO:"))}".Replace("OBSERVAÇÃO:", "").Replace("\n", " ");    
    linha++;
}


workbookHabitese.SaveAs($"{path}\\resultado\\ExcelHabitese.xlsx");

#endregion

#region alvara
/*
WorkBook workbookAlvara = WorkBook.Create(ExcelFileFormat.XLSX);
var sheetAlvara = workbookAlvara.CreateWorkSheet("ALVARA");
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
*/
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
*/
#endregion