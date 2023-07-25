

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
            objeto.Texto = PdfTextExtractor.GetTextFromPage(leitor, i).ToString().ToUpper();
            list.Add(objeto);
        }
            
    }

}

#endregion
/*try
{*/
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
    sheetHabitese["L1"].Value = "ESPECIFICAÇÃO (CATEGORIA)";
    sheetHabitese["M1"].Value = "ÁREA (CATEGORIA)(DESTINACAO)/(TIPO DE OBRA)/(m²))";
    sheetHabitese["N1"].Value = "ÁREA RESULTANTE";
    sheetHabitese["O1"].Value = "ÁREA LIBERADA";
    sheetHabitese["P1"].Value = "AREA TOTAL";
    sheetHabitese["Q1"].Value = "OBSERVACAO";


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
        if (ProprietarioImovel.IndexOf("CNPJ:") != -1)
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

        }
        else if (ResponsavelExecucaoObra.IndexOf("CPF:") != -1)
        {
            sheetHabitese[$"H{linha}"].Value = $"{ResponsavelExecucaoObra.Substring(ResponsavelExecucaoObra.IndexOf("NOME:"), ResponsavelExecucaoObra.IndexOf("CPF:") - ResponsavelExecucaoObra.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
            sheetHabitese[$"I{linha}"].Value = $"{ResponsavelExecucaoObra.Substring(ResponsavelExecucaoObra.IndexOf("CPF:"), ResponsavelExecucaoObra.Length - ResponsavelExecucaoObra.IndexOf("CPF:"))}".Replace("CPF:", "").Replace("\n", " ");

        }


        sheetHabitese[$"J{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:"), row.Texto.IndexOf("CONFORME") - row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:"))}".Replace("RESPONSÁVEL TÉCNICO:", "").Replace("\n", " ");
        sheetHabitese[$"K{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("CONFORME"), row.Texto.IndexOf("ESPECIFICAÇÃO:") - row.Texto.IndexOf("CONFORME"))}".Replace("\n", " ");

        var especificacoes = $"{row.Texto.Substring(row.Texto.IndexOf("ESPECIFICAÇÃO:"), row.Texto.IndexOf("ÁREA TOTAL DA OBRA:") - row.Texto.IndexOf("ESPECIFICAÇÃO:"))}".Replace("\n", " ");
        sheetHabitese[$"L{linha}"].Value = $"{especificacoes.Substring(especificacoes.IndexOf("TIPO DE HABITE-SE:"), especificacoes.IndexOf("DADOS DA OBRA:") - especificacoes.IndexOf("TIPO DE HABITE-SE:"))}".Replace("TIPO DE HABITE-SE:", "").Replace("\n", " ");

        var areaPrincipais = $"{row.Texto.Substring(row.Texto.IndexOf("ÁREAS PRINCIPAIS"), row.Texto.IndexOf("ÁREA TOTAL DA OBRA:") - row.Texto.IndexOf("ÁREAS PRINCIPAIS"))}".Replace("\n", " ");
        sheetHabitese[$"M{linha}"].Value = $"{areaPrincipais.Substring(areaPrincipais.IndexOf("TIPO DE OBRA ÁREA (M²)"), areaPrincipais.IndexOf("ÁREA RESULTANTE") - areaPrincipais.IndexOf("TIPO DE OBRA ÁREA (M²)"))}".Replace("TIPO DE OBRA ÁREA (M²)", "").Replace("\n", " ");
        sheetHabitese[$"N{linha}"].Value = $"{areaPrincipais.Substring(areaPrincipais.IndexOf("ÁREA RESULTANTE"), areaPrincipais.IndexOf("ÁREA LIBERADA") - areaPrincipais.IndexOf("ÁREA RESULTANTE"))}".Replace("ÁREA RESULTANTE", "").Replace("\n", " ");        
        sheetHabitese[$"O{linha}"].Value = $"{areaPrincipais.Substring(areaPrincipais.IndexOf("ÁREA LIBERADA"), areaPrincipais.Length - areaPrincipais.IndexOf("ÁREA LIBERADA"))}".Replace("ÁREA LIBERADA", "").Replace("\n", " ");
        
        sheetHabitese[$"P{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("ÁREA TOTAL DA OBRA:"), row.Texto.IndexOf("OBSERVAÇÃO:") - row.Texto.IndexOf("ÁREA TOTAL DA OBRA:"))}".Replace("ÁREA TOTAL DA OBRA:", "").Replace("\n", " ");
        sheetHabitese[$"Q{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("OBSERVAÇÃO:"), row.Texto.IndexOf("MG, EM") - row.Texto.IndexOf("OBSERVAÇÃO:"))}".Replace("OBSERVAÇÃO:", "").Replace("\n", " ");
        linha++;
    }


    workbookHabitese.SaveAs($"{path}\\resultado\\ExcelHabitese.xlsx");

    #endregion

    #region alvara modelo tabela

    WorkBook workbookAlvaraTabela = WorkBook.Create(ExcelFileFormat.XLSX);
    var sheetAlvaraTabela = workbookAlvaraTabela.CreateWorkSheet("ALVARA");

    sheetAlvaraTabela["A1"].Value = "NOME ARQUIVO";
    sheetAlvaraTabela["B1"].Value = "TITULO";
    sheetAlvaraTabela["C1"].Value = "PROPRIETÁRIO (NOME)";
    sheetAlvaraTabela["D1"].Value = "PROPRIETÁRIO DO IMÓVEL (CPF/CNPJ)";
    sheetAlvaraTabela["E1"].Value = "AUTOR PROJETO (NOME)";
    sheetAlvaraTabela["F1"].Value = "AUTOR PROJETO (CREA)";
    sheetAlvaraTabela["G1"].Value = "RESPONSÁVEL TÉCNICO (NOME)";
    sheetAlvaraTabela["H1"].Value = "RESPONSÁVEL TÉCNICO (CREA)";
    sheetAlvaraTabela["I1"].Value = "CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA (NOME)";
    sheetAlvaraTabela["J1"].Value = "CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA (CPF/CNPJ)";
    sheetAlvaraTabela["K1"].Value = "DESCRICAO";

    sheetAlvaraTabela["L1"].Value = "AREAS PRINCIPAIS";
   
    sheetAlvaraTabela["N1"].Value = "ÁREA RESULTANTE";
    sheetAlvaraTabela["M1"].Value = "ÁREA LIBERADA";
    sheetAlvaraTabela["O1"].Value = "OBSERVACAO";
    sheetAlvaraTabela["P1"].Value = "LEI";

    linha = 2;
    var listaAlvaraTabela = list.Where(tipo => tipo.Texto.Contains("CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA:") && 
    !tipo.Texto.Contains("//") && tipo.Texto.Contains("Áreas principais")).ToList();
    foreach (var row in listaAlvaraTabela)
    {
        sheetAlvaraTabela[$"A{linha}"].Value = $"{row.Nome}";
        sheetAlvaraTabela[$"B{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("ALVARÁ DE CONSTRUÇÃO Nº"), row.Texto.IndexOf("PROPRIETÁRIO:") - row.Texto.IndexOf("ALVARÁ DE CONSTRUÇÃO Nº"))}".Replace("\n", " ");

        var Proprietario = $"{row.Texto.Substring(row.Texto.IndexOf("PROPRIETÁRIO:"), row.Texto.IndexOf("AUTOR DO PROJETO:") - row.Texto.IndexOf("PROPRIETÁRIO:"))}";
        sheetAlvaraTabela[$"C{linha}"].Value = $"{Proprietario.Substring(Proprietario.IndexOf("NOME:"), Proprietario.IndexOf("CPF/CNPJ:") - Proprietario.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetAlvaraTabela[$"D{linha}"].Value = $"{Proprietario.Substring(Proprietario.IndexOf("CPF/CNPJ:"), Proprietario.Length - Proprietario.IndexOf("CPF/CNPJ:"))}".Replace("CPF/CNPJ:", "").Replace("\n", " ");

        var AutorProjeto = $"{row.Texto.Substring(row.Texto.IndexOf("AUTOR DO PROJETO:"), row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:") - row.Texto.IndexOf("AUTOR DO PROJETO:"))}";
        sheetAlvaraTabela[$"E{linha}"].Value = $"{AutorProjeto.Substring(AutorProjeto.IndexOf("NOME:"), AutorProjeto.IndexOf("CREA (CAU) Nº:") - AutorProjeto.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetAlvaraTabela[$"F{linha}"].Value = $"{AutorProjeto.Substring(AutorProjeto.IndexOf("CREA (CAU) Nº:"), AutorProjeto.Length - AutorProjeto.IndexOf("CREA (CAU) Nº:"))}".Replace("CREA (CAU) Nº:", "").Replace("\n", " ");

        var ResponsavelTecnico = $"{row.Texto.Substring(row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:"), row.Texto.IndexOf("CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA:") - row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:"))}";
        sheetAlvaraTabela[$"G{linha}"].Value = $"{ResponsavelTecnico.Substring(ResponsavelTecnico.IndexOf("NOME:"), ResponsavelTecnico.IndexOf("CREA (CAU) Nº:") - ResponsavelTecnico.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetAlvaraTabela[$"H{linha}"].Value = $"{ResponsavelTecnico.Substring(ResponsavelTecnico.IndexOf("CREA (CAU) Nº:"), ResponsavelTecnico.Length - ResponsavelTecnico.IndexOf("CREA (CAU) Nº:"))}".Replace("CREA (CAU) Nº:", "").Replace("\n", " ");

        var ResponsavelObra = $"{row.Texto.Substring(row.Texto.IndexOf("CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA:"), row.Texto.IndexOf("TENDO EM VISTA") - row.Texto.IndexOf("CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA:"))}";
        sheetAlvaraTabela[$"I{linha}"].Value = $"{ResponsavelObra.Substring(ResponsavelObra.IndexOf("NOME:"), ResponsavelObra.IndexOf("CPF/CNPJ:") - ResponsavelObra.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetAlvaraTabela[$"J{linha}"].Value = $"{ResponsavelObra.Substring(ResponsavelObra.IndexOf("CPF/CNPJ:"), ResponsavelObra.Length - ResponsavelObra.IndexOf("CPF/CNPJ:"))}".Replace("CPF/CNPJ:", "").Replace("\n", " ");

        sheetAlvaraTabela[$"K{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("TENDO EM VISTA"), row.Texto.IndexOf("Áreas principais") - row.Texto.IndexOf("TENDO EM VISTA"))}".Replace("\n", " ");

        var especificacaoIndice = -1;
        if (row.Texto.IndexOf("ESPECIFICAÇÃO") != -1)
        {
            especificacaoIndice = row.Texto.IndexOf("ESPECIFICAÇÃO");
        }else if (row.Texto.IndexOf("ESPECIFICAÇÃO") != -1)
        {
            especificacaoIndice = row.Texto.IndexOf("ESPECIFICAÇÃO");
        }

        var especificacoes = $"{row.Texto.Substring(row.Texto.IndexOf("Áreas principais"), especificacaoIndice - row.Texto.IndexOf("Áreas principais"))}".Replace("\n", " ");

        sheetAlvaraTabela[$"L{linha}"].Value = $"{especificacoes.Substring(especificacoes.IndexOf("TIPO DE OBRA ÁREA (M²)"), especificacoes.IndexOf("ÁREA RESULTANTE") - especificacoes.IndexOf("TIPO DE OBRA ÁREA (M²)"))}".Replace("TIPO DE OBRA ÁREA (M²)", "").Replace("\n", " ");
        
        sheetAlvaraTabela[$"M{linha}"].Value = $"{especificacoes.Substring(especificacoes.IndexOf("ÁREA RESULTANTE"), especificacoes.IndexOf("ÁREA LIBERADA") - especificacoes.IndexOf("ÁREA RESULTANTE"))}".Replace("ÁREA RESULTANTE", "").Replace("\n", " ");
        sheetAlvaraTabela[$"N{linha}"].Value = $"{especificacoes.Substring(especificacoes.IndexOf("ÁREA LIBERADA"), especificacoes.Length - especificacoes.IndexOf("ÁREA LIBERADA"))}".Replace("ÁREA LIBERADA", "").Replace("\n", " ");

        sheetAlvaraTabela[$"O{linha}"].Value = $"{row.Texto.Substring(especificacaoIndice, row.Texto.IndexOf("OBSERVAÇÕES:") - especificacaoIndice)}".Replace("ESPECIFICAÇÃO:", "").Replace("ESPECIFICAÇÃO:", "").Replace("\n", " ");
        sheetAlvaraTabela[$"P{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("OBSERVAÇÕES:"), row.Texto.IndexOf("LEI Nº") - row.Texto.IndexOf("OBSERVAÇÕES:"))}".Replace("OBSERVAÇÕES:", "").Replace("\n", " ");
        linha++;
    }
                               
    workbookAlvaraTabela.SaveAs($"{path}\\resultado\\ExcelAlvaraComTabela.xlsx");

    #endregion

    #region alvara modelo com barra

    WorkBook workbookAlvara = WorkBook.Create(ExcelFileFormat.XLSX);
    var sheetAlvaraBarra = workbookAlvara.CreateWorkSheet("ALVARA");

    sheetAlvaraBarra["A1"].Value = "NOME ARQUIVO";
    sheetAlvaraBarra["B1"].Value = "TITULO";
    sheetAlvaraBarra["C1"].Value = "PROPRIETÁRIO (NOME)";
    sheetAlvaraBarra["D1"].Value = "PROPRIETÁRIO DO IMÓVEL (CPF/CNPJ)";
    sheetAlvaraBarra["E1"].Value = "AUTOR PROJETO (NOME)";
    sheetAlvaraBarra["F1"].Value = "AUTOR PROJETO (CREA)";
    sheetAlvaraBarra["G1"].Value = "RESPONSÁVEL TÉCNICO (NOME)";
    sheetAlvaraBarra["H1"].Value = "RESPONSÁVEL TÉCNICO (CREA)";
    sheetAlvaraBarra["I1"].Value = "CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA (NOME)";
    sheetAlvaraBarra["J1"].Value = "CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA (CREA OU CNPJ/CPF)";
    sheetAlvaraBarra["K1"].Value = "DESCRICAO";
    sheetAlvaraBarra["L1"].Value = "ESPECIFICAÇÃO (DETALHES)";
    //sheetAlvaraBarra["M1"].Value = "AREA TOTAL DA OBRA";
    sheetAlvaraBarra["M1"].Value = "OBSERVACAO";
    sheetAlvaraBarra["N1"].Value = "LEI";

    linha = 2;
    var listaAlvara = list.Where(tipo => !tipo.Texto.Contains("ENDEREÇO DA OBRA:") && tipo.Texto.Contains("//")).ToList();
    foreach (var row in listaAlvara)
    {
        sheetAlvaraBarra[$"A{linha}"].Value = $"{row.Nome}";
        sheetAlvaraBarra[$"B{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("ALVARÁ DE CONSTRUÇÃO Nº"), row.Texto.IndexOf("PROPRIETÁRIO") - row.Texto.IndexOf("ALVARÁ DE CONSTRUÇÃO Nº"))}".Replace("\n", " ");

        var Proprietario = $"{row.Texto.Substring(row.Texto.IndexOf("PROPRIETÁRIO"), row.Texto.IndexOf("AUTOR DO PROJETO:") - row.Texto.IndexOf("PROPRIETÁRIO"))}";
        sheetAlvaraBarra[$"C{linha}"].Value = $"{Proprietario.Substring(Proprietario.IndexOf("NOME:"), Proprietario.IndexOf("CPF/CNPJ:") - Proprietario.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetAlvaraBarra[$"D{linha}"].Value = $"{Proprietario.Substring(Proprietario.IndexOf("CPF/CNPJ:"), Proprietario.Length - Proprietario.IndexOf("CPF/CNPJ:"))}".Replace("CPF/CNPJ:", "").Replace("\n", " ");

        var AutorProjeto = $"{row.Texto.Substring(row.Texto.IndexOf("AUTOR DO PROJETO:"), row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:") - row.Texto.IndexOf("AUTOR DO PROJETO:"))}";
        sheetAlvaraBarra[$"E{linha}"].Value = $"{AutorProjeto.Substring(AutorProjeto.IndexOf("NOME:"), AutorProjeto.IndexOf("CREA (CAU) Nº:") - AutorProjeto.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetAlvaraBarra[$"F{linha}"].Value = $"{AutorProjeto.Substring(AutorProjeto.IndexOf("CREA (CAU) Nº:"), AutorProjeto.Length - AutorProjeto.IndexOf("CREA (CAU) Nº:"))}".Replace("CREA (CAU) Nº:", "").Replace("\n", " ");

        var ResponsavelTecnico = $"{row.Texto.Substring(row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:"), row.Texto.IndexOf("CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA:") - row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:"))}";
        
        if (ResponsavelTecnico.IndexOf("CREA (CAU) Nº:") != -1)
        {
            sheetAlvaraBarra[$"G{linha}"].Value = $"{ResponsavelTecnico.Substring(ResponsavelTecnico.IndexOf("NOME:"), ResponsavelTecnico.IndexOf("CREA (CAU) Nº:") - ResponsavelTecnico.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
            sheetAlvaraBarra[$"H{linha}"].Value = $"{ResponsavelTecnico.Substring(ResponsavelTecnico.IndexOf("CREA (CAU) Nº:"), ResponsavelTecnico.Length - ResponsavelTecnico.IndexOf("CREA (CAU) Nº:"))}".Replace("CREA (CAU) Nº:", "").Replace("\n", " ");
        }
        else if (ResponsavelTecnico.IndexOf("CPF/CNPJ:") != -1)
        {
            sheetAlvaraBarra[$"G{linha}"].Value = $"{ResponsavelTecnico.Substring(ResponsavelTecnico.IndexOf("NOME:"), ResponsavelTecnico.IndexOf("CPF/CNPJ:") - ResponsavelTecnico.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
            sheetAlvaraBarra[$"H{linha}"].Value = $"{ResponsavelTecnico.Substring(ResponsavelTecnico.IndexOf("CPF/CNPJ:"), ResponsavelTecnico.Length - ResponsavelTecnico.IndexOf("CPF/CNPJ:"))}".Replace("CPF/CNPJ:", "").Replace("\n", " ");
        }

        var ConstrutoriaExecucao = $"{row.Texto.Substring(row.Texto.IndexOf("CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA:"), row.Texto.IndexOf("TENDO EM") - row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:"))}";
        
        if (ConstrutoriaExecucao.IndexOf("CREA (CAU) Nº:") != -1)
        {
            sheetAlvaraBarra[$"I{linha}"].Value = $"{ConstrutoriaExecucao.Substring(ConstrutoriaExecucao.IndexOf("NOME:"), ConstrutoriaExecucao.IndexOf("CREA (CAU) Nº:") - ConstrutoriaExecucao.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
            sheetAlvaraBarra[$"J{linha}"].Value = $"{ConstrutoriaExecucao.Substring(ConstrutoriaExecucao.IndexOf("CREA (CAU) Nº:"), ConstrutoriaExecucao.Length - ConstrutoriaExecucao.IndexOf("CREA (CAU) Nº:"))}".Replace("CREA (CAU) Nº:", "").Replace("\n", " ");
        }
        else if (ConstrutoriaExecucao.IndexOf("CPF/CNPJ:") != -1)
        {
            sheetAlvaraBarra[$"I{linha}"].Value = $"{ConstrutoriaExecucao.Substring(ConstrutoriaExecucao.IndexOf("NOME:"), ConstrutoriaExecucao.IndexOf("CPF/CNPJ:") - ConstrutoriaExecucao.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
            sheetAlvaraBarra[$"J{linha}"].Value = $"{ConstrutoriaExecucao.Substring(ConstrutoriaExecucao.IndexOf("CPF/CNPJ:"), ConstrutoriaExecucao.Length - ConstrutoriaExecucao.IndexOf("CPF/CNPJ:"))}".Replace("CPF/CNPJ:", "").Replace("\n", " ");
        }

        var especificacaoIndice = -1;
        if (row.Texto.IndexOf("ESPECIFICAÇÃO") != -1)
        {
            especificacaoIndice = row.Texto.IndexOf("ESPECIFICAÇÃO");
        }
        else if (row.Texto.IndexOf("ESPECIFICAÇÃO") != -1)
        {
            especificacaoIndice = row.Texto.IndexOf("ESPECIFICAÇÃO");
        }

        sheetAlvaraBarra[$"K{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("TENDO EM VISTA"), especificacaoIndice - row.Texto.IndexOf("TENDO EM VISTA"))}".Replace("\n", " ");

        var especificacoes = $"{row.Texto.Substring(especificacaoIndice, row.Texto.IndexOf("OBSERVAÇÕES:") - especificacaoIndice)}".Replace("\n", " ");

        sheetAlvaraBarra[$"L{linha}"].Value = $"{especificacoes.Substring(0, especificacoes.Length - 0)}".Replace("\n", " ");        
        //sheetAlvaraBarra[$"M{linha}"].Value = $"{especificacoes.Substring(especificacoes.IndexOf("ÁREA RESULTANTE"), especificacoes.IndexOf("ÁREA LIBERADA") - especificacoes.IndexOf("ÁREA RESULTANTE"))}".Replace("ÁREA RESULTANTE", "").Replace("\n", " ");        
        //sheetAlvaraBarra[$"M{linha}"].Value = $"{especificacoes.Substring(especificacoes.IndexOf("AREA TOTAL DA OBRA"), especificacoes.Length - especificacoes.IndexOf("AREA TOTAL DA OBRA"))}".Replace("AREA TOTAL DA OBRA", "").Replace("\n", " ");

        sheetAlvaraBarra[$"M{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("OBSERVAÇÕES:"), row.Texto.IndexOf("LEI Nº") - row.Texto.IndexOf("OBSERVAÇÕES:"))}".Replace("OBSERVAÇÕES:", "").Replace("\n", " ");
        sheetAlvaraBarra[$"N{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("LEI Nº"), row.Texto.IndexOf("ALFENAS - MG, EM") - row.Texto.IndexOf("LEI Nº"))}".Replace("\n", " ");
        linha++;
    }

    workbookAlvara.SaveAs($"{path}\\resultado\\ExcelAlvaraComBarra.xlsx");

    #endregion
/*
    #region alvara sem especificacao

    WorkBook workbookAlvaraSemEspecificacao = WorkBook.Create(ExcelFileFormat.XLSX);
    var sheetAlvaraBarraSemEspecificacao = workbookAlvara.CreateWorkSheet("ALVARA");

    sheetAlvaraBarraSemEspecificacao["A1"].Value = "NOME ARQUIVO";
    sheetAlvaraBarraSemEspecificacao["B1"].Value = "TITULO";
    sheetAlvaraBarraSemEspecificacao["C1"].Value = "PROPRIETÁRIO (NOME)";
    sheetAlvaraBarraSemEspecificacao["D1"].Value = "PROPRIETÁRIO DO IMÓVEL (CPF/CNPJ)";
    sheetAlvaraBarraSemEspecificacao["E1"].Value = "AUTOR PROJETO (NOME)";
    sheetAlvaraBarraSemEspecificacao["F1"].Value = "AUTOR PROJETO (CREA)";
    sheetAlvaraBarraSemEspecificacao["G1"].Value = "RESPONSÁVEL TÉCNICO (NOME)";
    sheetAlvaraBarraSemEspecificacao["H1"].Value = "RESPONSÁVEL TÉCNICO (CREA)";
    sheetAlvaraBarraSemEspecificacao["I1"].Value = "CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA (NOME)";
    sheetAlvaraBarraSemEspecificacao["J1"].Value = "CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA (CREA OU CNPJ/CPF)";
    sheetAlvaraBarraSemEspecificacao["K1"].Value = "DESCRICAO";
    sheetAlvaraBarraSemEspecificacao["L1"].Value = "AREAS PRINCIPAIS";
    sheetAlvaraBarraSemEspecificacao["M1"].Value = "ESPECIFICAÇÃO (DETALHES)";
    sheetAlvaraBarraSemEspecificacao["N1"].Value = "OBSERVACAO";
    sheetAlvaraBarraSemEspecificacao["O1"].Value = "LEI";

    linha = 2;
    var listaAlvaraSemEspecificacao = list.Where(tipo => tipo.Texto.Contains("//") && !tipo.Texto.Contains("Áreas principais")).ToList();
    foreach (var row in listaAlvaraSemEspecificacao)
    {
        sheetAlvaraBarraSemEspecificacao[$"A{linha}"].Value = $"{row.Nome}";
        sheetAlvaraBarraSemEspecificacao[$"B{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("ALVARÁ DE CONSTRUÇÃO Nº"), row.Texto.IndexOf("PROPRIETÁRIO:") - row.Texto.IndexOf("ALVARÁ DE CONSTRUÇÃO Nº"))}".Replace("\n", " ");

        var Proprietario = $"{row.Texto.Substring(row.Texto.IndexOf("PROPRIETÁRIO:"), row.Texto.IndexOf("AUTOR DO PROJETO:") - row.Texto.IndexOf("PROPRIETÁRIO:"))}";
        sheetAlvaraBarraSemEspecificacao[$"C{linha}"].Value = $"{Proprietario.Substring(Proprietario.IndexOf("NOME:"), Proprietario.IndexOf("CPF/CNPJ:") - Proprietario.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetAlvaraBarraSemEspecificacao[$"D{linha}"].Value = $"{Proprietario.Substring(Proprietario.IndexOf("CPF/CNPJ:"), Proprietario.Length - Proprietario.IndexOf("CPF/CNPJ:"))}".Replace("CPF/CNPJ:", "").Replace("\n", " ");

        var AutorProjeto = $"{row.Texto.Substring(row.Texto.IndexOf("AUTOR DO PROJETO:"), row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:") - row.Texto.IndexOf("AUTOR DO PROJETO:"))}";
        sheetAlvaraBarraSemEspecificacao[$"E{linha}"].Value = $"{AutorProjeto.Substring(AutorProjeto.IndexOf("NOME:"), AutorProjeto.IndexOf("CREA (CAU) Nº:") - AutorProjeto.IndexOf("NOME:"))}".Replace("NOME:", "").Replace("\n", " ");
        sheetAlvaraBarraSemEspecificacao[$"F{linha}"].Value = $"{AutorProjeto.Substring(AutorProjeto.IndexOf("CREA (CAU) Nº:"), AutorProjeto.Length - AutorProjeto.IndexOf("CREA (CAU) Nº:"))}".Replace("CREA (CAU) Nº:", "").Replace("\n", " ");

        var ConstrutoriaExecucao = $"{row.Texto.Substring(row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:"), row.Texto.IndexOf("CONSTRUTORA OU RESPONSÁVEL PELA EXECUÇÃO DA OBRA:") - row.Texto.IndexOf("RESPONSÁVEL TÉCNICO:"))}";
        if (ConstrutoriaExecucao.IndexOf("CREA (CAU) Nº:") != -1)
        {
            sheetAlvaraBarraSemEspecificacao[$"G{linha}"].Value = $"{ConstrutoriaExecucao.Substring(ConstrutoriaExecucao.IndexOf("CREA (CAU) Nº:"), ConstrutoriaExecucao.Length - ConstrutoriaExecucao.IndexOf("CREA (CAU) Nº:"))}".Replace("CREA (CAU) Nº:", "").Replace("\n", " ");
        }
        else if (ConstrutoriaExecucao.IndexOf("CPF/CNPJ:") != -1)
        {
            sheetAlvaraBarraSemEspecificacao[$"G{linha}"].Value = $"{ConstrutoriaExecucao.Substring(ConstrutoriaExecucao.IndexOf("CPF/CNPJ:"), ConstrutoriaExecucao.Length - ConstrutoriaExecucao.IndexOf("CPF/CNPJ:"))}".Replace("CPF/CNPJ:", "").Replace("\n", " ");
        }

        var especificacaoIndice = -1;
        if (row.Texto.IndexOf("ESPECIFICAÇÃO") != -1)
        {
            especificacaoIndice = row.Texto.IndexOf("ESPECIFICAÇÃO");
        }
        else if (row.Texto.IndexOf("ESPECIFICAÇÃO") != -1)
        {
            especificacaoIndice = row.Texto.IndexOf("ESPECIFICAÇÃO");
        }

        sheetAlvaraBarraSemEspecificacao[$"H{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("TENDO EM VISTA"), especificacaoIndice - row.Texto.IndexOf("TENDO EM VISTA"))}".Replace("\n", " ");

        var especificacoes = $"{row.Texto.Substring(row.Texto.IndexOf("Áreas principais"), especificacaoIndice - row.Texto.IndexOf("Áreas principais"))}".Replace("\n", " ");

        sheetAlvaraBarraSemEspecificacao[$"L{linha}"].Value = $"{especificacoes.Substring(especificacoes.IndexOf("TIPO DE OBRA ÁREA (M²)"), especificacoes.IndexOf("ÁREA RESULTANTE") - especificacoes.IndexOf("TIPO DE OBRA ÁREA (M²)"))}".Replace("TIPO DE OBRA ÁREA (M²)", "").Replace("\n", " ");
        sheetAlvaraBarraSemEspecificacao[$"M{linha}"].Value = $"{especificacoes.Substring(especificacoes.IndexOf("ÁREA RESULTANTE"), especificacoes.IndexOf("ÁREA LIBERADA") - especificacoes.IndexOf("ÁREA RESULTANTE"))}".Replace("ÁREA RESULTANTE", "").Replace("\n", " ");
        sheetAlvaraBarraSemEspecificacao[$"N{linha}"].Value = $"{especificacoes.Substring(especificacoes.IndexOf("ÁREA LIBERADA"), especificacoes.Length - especificacoes.IndexOf("ÁREA LIBERADA"))}".Replace("ÁREA LIBERADA", "").Replace("\n", " ");

        sheetAlvaraBarraSemEspecificacao[$"O{linha}"].Value = $"{row.Texto.Substring(especificacaoIndice, row.Texto.IndexOf("OBSERVAÇÕES:") - especificacaoIndice)}".Replace("ESPECIFICAÇÃO:", "").Replace("ESPECIFICAÇÃO:", "").Replace("\n", " ");
        sheetAlvaraBarraSemEspecificacao[$"P{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("OBSERVAÇÕES:"), row.Texto.IndexOf("LEI Nº") - row.Texto.IndexOf("OBSERVAÇÕES:"))}".Replace("OBSERVAÇÕES:", "").Replace("\n", " ");
        sheetAlvaraBarraSemEspecificacao[$"P{linha}"].Value = $"{row.Texto.Substring(row.Texto.IndexOf("LEI Nº 856 DE 23/11/1964"), row.Texto.IndexOf("ALFENAS - MG, EM") - row.Texto.IndexOf("LEI Nº 856 DE 23/11/1964"))}".Replace("\n", " ");
        linha++;
    }

    workbookAlvaraSemEspecificacao.SaveAs($"{path}\\resultado\\ExcelAlvara.xlsx");

    #endregion*/
/*}
catch (Exception e)
{
    System.Console.WriteLine("ERRO:" + e.Message);
    throw;
}*/



