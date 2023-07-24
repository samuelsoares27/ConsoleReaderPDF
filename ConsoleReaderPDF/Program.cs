

using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Reflection;
using System.Text;


List<string> list = new ();
bool habiteses = false;
bool alvara = false;
var contador = 1;

string[] arquivos = Directory.GetFiles("C:\\Trabalho\\Arquivos\\prefeitura\\alvarahabilite");

Console.WriteLine("Arquivos:");
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