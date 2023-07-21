

using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Text;


List<string> list = new ();

bool habiteses = false;
bool alvara = false;

using (PdfReader leitor = new ("C:\\temp\\Habitese - 2021-04-20T151854.374.pdf"))
{
    StringBuilder texto = new ();
    for (int i = 1; i <= leitor.NumberOfPages; i++)
    {
        list.Add(PdfTextExtractor.GetTextFromPage(leitor, i));

        if (list[i-1].Contains("HABITE-SE"))
        {
            habiteses = true;
        }
        if (list[i-1].Contains("ALVARÁ"))
        {
            alvara = true;
        }
    }

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

