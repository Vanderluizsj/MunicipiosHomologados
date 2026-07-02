using System.Reflection;
using MunicipiosHomologados.ConsoleApp.Util;
using OfficeOpenXml;

namespace MunicipiosHomologados.ConsoleApp.Services;

public class IbgeService
{
    private const string NomeBaseIbge = "Base_Consulta_IBGE_Municipios_2024.xlsx";
    public Dictionary<string, string> CarregarMunicipios()
    {
        Logger.Info("📂 Carregando base de dados interna do IBGE...");

        return CarregarDicionarioIbgeEmbutido();
    }

    private Dictionary<string, string> CarregarDicionarioIbgeEmbutido()    
    {
        var dicionario = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var assembly = Assembly.GetExecutingAssembly();

        string nomeRecurso = "";

        foreach (var name in assembly.GetManifestResourceNames())
        {
            if (name.Contains(NomeBaseIbge))
            {
                nomeRecurso = name;
                break;
            }
        }

        if (string.IsNullOrEmpty(nomeRecurso))
            throw new Exception("Base do IBGE não encontrada.");

        using Stream? stream = assembly.GetManifestResourceStream(nomeRecurso);

        if (stream == null)
            throw new Exception("Falha ao abrir a planilha.");

        using var pacote = new ExcelPackage(stream);

        var ws = pacote.Workbook.Worksheets[0];

        int totalLinhas = ws.Dimension?.End.Row ?? 0;

        for (int linha = 2; linha <= totalLinhas; linha++)
        {
            string chave = ws.Cells[linha, 1].Text.Trim().ToUpper();
            string codigo = ws.Cells[linha, 2].Text.Trim();

            if (!string.IsNullOrEmpty(chave) &&
                !string.IsNullOrEmpty(codigo))
            {
                dicionario.TryAdd(chave, codigo);
            }
        }

        return dicionario;
    }
}