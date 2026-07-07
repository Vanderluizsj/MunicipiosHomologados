using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using MunicipiosHomologados.ConsoleApp.Models;
using MunicipiosHomologados.ConsoleApp.Util;
using OfficeOpenXml;

namespace MunicipiosHomologados.ConsoleApp.Services
{
    public class IbgeService
    {
        private const string NomeBaseIbge = "Base_Consulta_IBGE_Municipios_2024.xlsx";

        public Dictionary<string, MunicipioIbge> CarregarMunicipios()
        {
            Logger.Info("📂 Carregando base de dados interna do IBGE (Embarcada)...");

            var dicionario = new Dictionary<string, MunicipioIbge>(StringComparer.OrdinalIgnoreCase);
            var assembly = Assembly.GetExecutingAssembly();
            string nomeRecurso = "";

            // Localiza o arquivo embutido no manifesto do assembly
            foreach (var name in assembly.GetManifestResourceNames())
            {
                if (name.Contains(NomeBaseIbge))
                {
                    nomeRecurso = name;
                    break;
                }
            }

            if (string.IsNullOrEmpty(nomeRecurso))
                throw new Exception($"Base do IBGE '{NomeBaseIbge}' não foi encontrada como Recurso Embutido.");

            using (Stream stream = assembly.GetManifestResourceStream(nomeRecurso))
            {
                if (stream == null)
                    throw new Exception("Falha ao abrir o fluxo da planilha embutida do IBGE.");

                ExcelPackage.License.SetNonCommercialPersonal("Uso Pessoal");

                using (var pacote = new ExcelPackage(stream))
                {
                    var ws = pacote.Workbook.Worksheets[0];
                    int totalLinhas = ws.Dimension?.End.Row ?? 0;

                    for (int linha = 2; linha <= totalLinhas; linha++)
                    {
                        string chave = ws.Cells[linha, 1].Text.Trim().ToUpper(); // Ex: "SÃO PAULO - SP"
                        string codigo = ws.Cells[linha, 2].Text.Trim();          // Ex: "3550308"

                        if (!string.IsNullOrEmpty(chave) && !string.IsNullOrEmpty(codigo))
                        {
                            var municipio = new MunicipioIbge
                            {
                                ChavePadrao = chave,
                                CodigoFormatado = codigo
                            };

                            dicionario.TryAdd(chave, municipio);
                        }
                    }
                }
            }

            return dicionario;
        }
    }
}