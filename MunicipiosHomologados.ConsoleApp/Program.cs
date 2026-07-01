using System.Drawing; // Requerido para gerenciar as cores do Excel
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style; // Requerido para aplicar as bordas e alinhamentos

namespace ValidadorEscalavelNDD
{
    class Program
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        static async Task Main(string[] args)
        {
            // CAPTURA A PASTA ONDE O EXECUTÁVEL ESTÁ RODANDO ATUALMENTE
            string pastaExecutavel = AppDomain.CurrentDomain.BaseDirectory;

            // CAMINHOS DE ENTRADA (O .NET vai garantir que elas existam na mesma pasta do .exe)
            string caminhoBaseIbge = Path.Combine(pastaExecutavel, "Base_Consulta_IBGE_Municipios_2024.xlsx");
            string caminhoClientes = Path.Combine(pastaExecutavel, "Cliente.xlsx");

            // CAMINHO DE SAÍDA (O resultado final sempre será gerado na pasta atual do .exe)
            string caminhoResultado = Path.Combine(pastaExecutavel, "Cliente_Resultado_Final.xlsx");
            string urlNdd = "https://documentacao-nfse.e-datacenter.nddigital.com.br/fiscal-documentacao/docs/ndd-nfse/municipios-nfse/";
            // VALIDAÇÃO DA LICENÇA INDIVIDUAL DO EPPLUS 8+
            ExcelPackage.License.SetNonCommercialPersonal("<Vander>");

            try
            {
                Console.WriteLine("🚀 Iniciando processamento escalável e estilização visual...");

                // 1. CARREGA O SITE DA NDD EXTRAINDO OS CÓDIGOS E DETECTANDO O PADRÃO NACIONAL
                Console.WriteLine("🌐 Consultando site da NDD...");
                var municipiosHomologadosNdd = await ObterMunicipiosHomologadosNddAsync(urlNdd);

                // 2. CARREGA A BASE DO IBGE EM MEMÓRIA
                Console.WriteLine("📂 Carregando base de dados do IBGE para o PROCV interno...");
                var dicionarioIbge = CarregarDicionarioIbgeLocal(caminhoBaseIbge);

                // 3. PROCESSA A PLANILHA DE CLIENTES
                Console.WriteLine("📝 Processando dados dos clientes...");

                if (!File.Exists(caminhoClientes))
                {
                    throw new FileNotFoundException($"O arquivo de clientes não foi encontrado: {caminhoClientes}");
                }

                using (var pacote = new ExcelPackage(new FileInfo(caminhoClientes)))
                {
                    if (pacote.Workbook.Worksheets.Count == 0)
                    {
                        throw new Exception($"Nenhuma aba legível encontrada em '{caminhoClientes}'.");
                    }

                    var ws = pacote.Workbook.Worksheets[0];
                    int totalLinhas = ws.Dimension?.End.Row ?? 0;

                    // DEFINE OS CABEÇALHOS DAS COLUNAS C, D E E
                    ws.Cells[1, 3].Value = "Código IBGE";
                    ws.Cells[1, 4].Value = "Status NDD";
                    ws.Cells[1, 5].Value = "NFSe Nacional";

                    // Habilita a visualização das linhas de grade padrão do Excel
                    ws.View.ShowGridLines = true;

                    for (int linha = 2; linha <= totalLinhas; linha++)
                    {
                        string municipio = ws.Cells[linha, 1].Value?.ToString()?.Trim() ?? "";
                        string uf = ws.Cells[linha, 2].Value?.ToString()?.Trim() ?? "";

                        if (string.IsNullOrEmpty(municipio) || string.IsNullOrEmpty(uf)) continue;

                        string chaveProcv = $"{municipio} - {uf}".ToUpper();

                        if (dicionarioIbge.TryGetValue(chaveProcv, out string? codigoIbge) && !string.IsNullOrEmpty(codigoIbge))
                        {
                            if (long.TryParse(codigoIbge, out long codigoNum))
                            {
                                ws.Cells[linha, 3].Value = codigoNum;
                                ws.Cells[linha, 3].Style.Numberformat.Format = "0000000";
                            }
                            else
                            {
                                ws.Cells[linha, 3].Value = codigoIbge;
                            }

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

                    // =========================================================================
                    // BLOCO DE FORMATAÇÃO VISUAL CORRIGIDO PARA O EPPLUS 8
                    // =========================================================================
                    Console.WriteLine("🎨 Aplicando design e estilização nas células...");

                    // 1. Estilização do Cabeçalho (Linha 1, Colunas A até E)
                    using (var rangeHeader = ws.Cells[1, 1, 1, 5])
                    {
                        rangeHeader.Style.Font.Name = "Arial";
                        rangeHeader.Style.Font.Size = 11;
                        rangeHeader.Style.Font.Bold = true;
                        rangeHeader.Style.Font.Color.SetColor(Color.White);

                        // Cor Azul Escuro (Navy Blue)
                        rangeHeader.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rangeHeader.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(31, 78, 120));

                        // CORREÇÃO: Propriedades diretas de alinhamento do EPPlus
                        rangeHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        rangeHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    // 2. Estilização das Linhas de Dados e Efeito Zebrado
                    Color corZebradaSuave = Color.FromArgb(249, 251, 253);
                    Color corBordaCinza = Color.FromArgb(217, 217, 217);

                    for (int r = 2; r <= totalLinhas; r++)
                    {
                        var linhaDados = ws.Cells[r, 1, r, 5];
                        linhaDados.Style.Font.Name = "Arial";
                        linhaDados.Style.Font.Size = 10;

                        // CORREÇÃO: Alinhamentos específicos diretos por coluna
                        ws.Cells[r, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;   // Município
                        ws.Cells[r, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // UF
                        ws.Cells[r, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Código IBGE
                        ws.Cells[r, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;   // Status
                        ws.Cells[r, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // NFSe Nacional

                        // Efeito Zebrado (Linhas pares)
                        if (r % 2 == 0)
                        {
                            linhaDados.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            linhaDados.Style.Fill.BackgroundColor.SetColor(corZebradaSuave);
                        }

                        // Aplica bordas finas em todas as células de dados
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

                    // 3. Ajuste Automático da Largura das Colunas
                    ws.Cells[1, 1, totalLinhas, 5].AutoFitColumns();

                    // Salva os resultados no arquivo final formatado
                    pacote.SaveAs(new FileInfo(caminhoResultado));
                    Console.WriteLine($"\n✔ Concluído! Planilha formatada gerada em: {caminhoResultado}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n❌ Erro crítico no processo: {ex.Message}");
            }
        }

        private static async Task<Dictionary<string, bool>> ObterMunicipiosHomologadosNddAsync(string url)
        {
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) C# Core");

            string html = await _httpClient.GetStringAsync(url);

            if (!html.Contains("Municípios Homologados", StringComparison.OrdinalIgnoreCase))
                throw new Exception("A página da NDD não retornou o conteúdo esperado.");

            var mapaMunicipios = new Dictionary<string, bool>();

            var regexLinhas = new Regex(@"([^\n\r]{1,60})\b([0-9]{7})\b([^\n\r]{1,60})", RegexOptions.IgnoreCase);
            var matches = regexLinhas.Matches(html);

            foreach (Match match in matches)
            {
                string textoAoRedor = match.Value;
                string codigoIbge = match.Groups[2].Value;

                bool ehNacional = textoAoRedor.Contains("Nacional", StringComparison.OrdinalIgnoreCase) ||
                                  textoAoRedor.Contains("NfseNacional", StringComparison.OrdinalIgnoreCase);

                if (!mapaMunicipios.ContainsKey(codigoIbge))
                {
                    mapaMunicipios.Add(codigoIbge, ehNacional);
                }
                else if (ehNacional)
                {
                    mapaMunicipios[codigoIbge] = true;
                }
            }

            return mapaMunicipios;
        }

        private static Dictionary<string, string> CarregarDicionarioIbgeLocal(string caminho)
        {
            var dicionario = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using (var pacote = new ExcelPackage(new FileInfo(caminho)))
            {
                var ws = pacote.Workbook.Worksheets[0];
                int totalLinhas = ws.Dimension?.End.Row ?? 0;

                for (int linha = 2; linha <= totalLinhas; linha++)
                {
                    string chaveOriginal = ws.Cells[linha, 1].Value?.ToString()?.Trim() ?? "";
                    string codigo = ws.Cells[linha, 2].Value?.ToString()?.Trim() ?? "";

                    if (!string.IsNullOrEmpty(chaveOriginal) && !string.IsNullOrEmpty(codigo))
                    {
                        string chave = chaveOriginal.ToUpper();

                        if (!dicionario.ContainsKey(chave))
                        {
                            dicionario.Add(chave, codigo);
                        }
                    }
                }
            }
            return dicionario;
        }
    }
}