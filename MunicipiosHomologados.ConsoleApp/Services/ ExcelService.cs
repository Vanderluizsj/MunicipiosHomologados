using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using MunicipiosHomologados.ConsoleApp.Util; // Caso queira usar seu Logger personalizado futuramente

namespace MunicipiosHomologados.ConsoleApp.Services
{
    public class ExcelService
    {
        public void ProcessarPlanilhaClientes(Dictionary<string, string> dicionarioIbge, Dictionary<string, bool> municipiosHomologadosNdd)
        {
            Logger.Info("📝 Processando e higienizando dados dos clientes...");

            // Caminhos baseados na execução do usuário (pasta atual do .exe)
            string pastaExecutavel = AppDomain.CurrentDomain.BaseDirectory;
            string caminhoClientes = Path.Combine(pastaExecutavel, "Cliente.xlsx");
            string caminhoResultado = Path.Combine(pastaExecutavel, "Cliente_Resultado_Final.xlsx");

            if (!File.Exists(caminhoClientes))
            {
                throw new FileNotFoundException($"O arquivo obrigatório 'Cliente.xlsx' não foi encontrado na pasta atual.");
            }            

            using (var pacote = new ExcelPackage(new FileInfo(caminhoClientes)))
            {
                if (pacote.Workbook.Worksheets.Count == 0)
                {
                    throw new Exception("Nenhuma aba legível encontrada em 'Cliente.xlsx'.");
                }

                var ws = pacote.Workbook.Worksheets[0];
                int totalLinhas = ws.Dimension?.End.Row ?? 0;

                if (totalLinhas < 2)
                {
                    Logger.Info("⚠ Nenhuma linha de dados encontrada abaixo do cabeçalho de clientes.");
                    return;
                }

                // Define os Cabeçalhos das novas colunas
                ws.Cells[1, 3].Value = "Código IBGE";
                ws.Cells[1, 4].Value = "Status NDD";
                ws.Cells[1, 5].Value = "NFSe Nacional";

                // Habilita as linhas de grade padrões do Excel
                ws.View.ShowGridLines = true;

                // Processamento dos dados linha por linha
                for (int linha = 2; linha <= totalLinhas; linha++)
                {
                    string municipio = ws.Cells[linha, 1].Value?.ToString()?.Trim() ?? "";
                    string uf = ws.Cells[linha, 2].Value?.ToString()?.Trim() ?? "";

                    if (string.IsNullOrEmpty(municipio) || string.IsNullOrEmpty(uf)) continue;

                    // Chave de cruzamento do PROCV
                    string chaveProcv = $"{municipio} - {uf}".ToUpper();

                    // EXECUTA O PROCV EM MEMÓRIA
                    if (dicionarioIbge.TryGetValue(chaveProcv, out string? codigoIbge) && !string.IsNullOrEmpty(codigoIbge))
                    {
                        // Coluna C: Grava o código IBGE com 7 dígitos
                        if (long.TryParse(codigoIbge, out long codigoNum))
                        {
                            ws.Cells[linha, 3].Value = codigoNum;
                            ws.Cells[linha, 3].Style.Numberformat.Format = "0000000";
                        }
                        else
                        {
                            ws.Cells[linha, 3].Value = codigoIbge;
                        }

                        // Colunas D e E: Valida contra a tabela processada do site da NDD
                        if (municipiosHomologadosNdd.TryGetValue(codigoIbge, out bool ehNacional))
                        {
                            ws.Cells[linha, 4].Value = "Homologado";
                            ws.Cells[linha, 5].Value = ehNacional ? "Sim" : "Não";
                        }
                        else
                        {
                            ws.Cells[linha, 4].Value = "Não Homologado";
                            ws.Cells[linha, 5].Value = "Não";
                        }
                    }
                    else
                    {
                        ws.Cells[linha, 3].Value = "-";
                        ws.Cells[linha, 4].Value = "Município não encontrado na Base";
                        ws.Cells[linha, 5].Value = "Não";
                    }
                }

                // Aplica a identidade visual do relatório
                AplicarLayoutProfissional(ws, totalLinhas);

                // Salva na raiz do executável
                pacote.SaveAs(new FileInfo(caminhoResultado));
                Logger.Info($"\n✔ Sucesso! Arquivo gerado em: {caminhoResultado}");
            }
        }

        private void AplicarLayoutProfissional(ExcelWorksheet ws, int totalLinhas)
        {
            Logger.Info("🎨 Formatando layout da tabela de resultados...");

            // 1. Estilização do Cabeçalho (Linha 1, Colunas A até E)
            using (var rangeHeader = ws.Cells[1, 1, 1, 5])
            {
                rangeHeader.Style.Font.Name = "Arial";
                rangeHeader.Style.Font.Size = 11;
                rangeHeader.Style.Font.Bold = true;
                rangeHeader.Style.Font.Color.SetColor(Color.White);
                rangeHeader.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rangeHeader.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(31, 78, 120)); // Azul Corporativo Navy
                rangeHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                rangeHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            // Paletas de cores para o design de dados
            Color corZebradaSuave = Color.FromArgb(249, 251, 253);
            Color corBordaCinza = Color.FromArgb(217, 217, 217);

            // 2. Estilização das Linhas de Dados e Alinhamentos
            for (int r = 2; r <= totalLinhas; r++)
            {
                var linhaDados = ws.Cells[r, 1, r, 5];
                linhaDados.Style.Font.Name = "Arial";
                linhaDados.Style.Font.Size = 10;

                ws.Cells[r, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;   // Município
                ws.Cells[r, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // UF
                ws.Cells[r, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Código IBGE
                ws.Cells[r, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;   // Status NDD
                ws.Cells[r, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // NFSe Nacional

                // Efeito Zebrado nas Linhas Pares
                if (r % 2 == 0)
                {
                    linhaDados.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    linhaDados.Style.Fill.BackgroundColor.SetColor(corZebradaSuave);
                }

                // Moldura de Bordas Finas Cinzas
                for (int col = 1; col <= 5; col++)
                {
                    var cell = ws.Cells[r, col];
                    cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Top.Color.SetColor(corBordaCinza);
                    cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Bottom.Color.SetColor(corBordaCinza);
                    cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Left.Color.SetColor(corBordaCinza);
                    cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Right.Color.SetColor(corBordaCinza);
                }
            }

            // 3. Ajuste de Dimensionamento de Coluna Automático
            ws.Cells[1, 1, totalLinhas, 5].AutoFitColumns();
        }
    }
}