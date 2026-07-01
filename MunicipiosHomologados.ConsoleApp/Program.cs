using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeOpenXml; // Requer o pacote NuGet: EPPlus 8+

namespace ValidadorEscalavelNDD
{
    class Program
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        static async Task Main(string[] args)
        {
            // CAMINHOS ABSOLUTOS DIRETOS PARA A SUA PASTA DE ORIGEM
            string pastaTrabalho = @"C:\Source\MunicipiosHomologados\";
            string caminhoBaseIbge = Path.Combine(pastaTrabalho, "Base_Consulta_IBGE_Municipios_2024.xlsx");
            string caminhoClientes = Path.Combine(pastaTrabalho, "Cliente.xlsx");
            string caminhoResultado = Path.Combine(pastaTrabalho, "Cliente_Resultado_Final.xlsx");
            string urlNdd = "https://documentacao-nfse.e-datacenter.nddigital.com.br/fiscal-documentacao/docs/ndd-nfse/municipios-nfse/";

            // VALIDAÇÃO DA LICENÇA INDIVIDUAL REQUERIDA NO EPPLUS 8+
            ExcelPackage.License.SetNonCommercialPersonal("<Vander>");

            try
            {
                Console.WriteLine("🚀 Iniciando processamento escalável...");

                // 1. CARREGA O SITE DA NDD EXTRAINDO OS CÓDIGOS E DETECTANDO O PADRÃO NACIONAL
                Console.WriteLine("🌐 Consultando site da NDD...");
                var municipiosHomologadosNdd = await ObterMunicipiosHomologadosNddAsync(urlNdd);

                // 2. CARREGA A BASE DO IBGE EM MEMÓRIA (O NOSSO "PROCV" EM C#)
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
                    ws.Cells[1, 5].Value = "NFSe Nacional"; // Nova Coluna E

                    int totalAtualizados = 0;

                    for (int linha = 2; linha <= totalLinhas; linha++)
                    {
                        string municipio = ws.Cells[linha, 1].Value?.ToString()?.Trim() ?? "";
                        string uf = ws.Cells[linha, 2].Value?.ToString()?.Trim() ?? "";

                        if (string.IsNullOrEmpty(municipio) || string.IsNullOrEmpty(uf)) continue;

                        // Monta a chave em maiúsculo para buscar no dicionário: "MUNICÍPIO - UF"
                        string chaveProcv = $"{municipio} - {uf}".ToUpper();

                        // EXECUTA O PROCV EM MEMÓRIA
                        if (dicionarioIbge.TryGetValue(chaveProcv, out string? codigoIbge) && !string.IsNullOrEmpty(codigoIbge))
                        {
                            // Grava o código do IBGE formatado com 7 dígitos na Coluna C
                            if (long.TryParse(codigoIbge, out long codigoNum))
                            {
                                ws.Cells[linha, 3].Value = codigoNum;
                                ws.Cells[linha, 3].Style.Numberformat.Format = "0000000";
                            }
                            else
                            {
                                ws.Cells[linha, 3].Value = codigoIbge;
                            }

                            // VALIDAÇÃO INTELIGENTE CONTRA O MAPA DA NDD
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
                            totalAtualizados++;
                        }
                        else
                        {
                            ws.Cells[linha, 3].Value = "-";
                            ws.Cells[linha, 4].Value = "Município não encontrado na Base";
                            ws.Cells[linha, 5].Value = "Não";
                        }
                    }

                    // Salva os resultados no arquivo final
                    pacote.SaveAs(new FileInfo(caminhoResultado));
                    Console.WriteLine($"\n✔ Concluído! {totalAtualizados} municípios validados.");
                    Console.WriteLine($"📊 Resultado exportado para: {caminhoResultado}");
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

            // Captura o código de 7 dígitos junto com os caracteres vizinhos para checar o contexto
            var regexLinhas = new Regex(@"([^\n\r]{1,60})\b([0-9]{7})\b([^\n\r]{1,60})", RegexOptions.IgnoreCase);
            var matches = regexLinhas.Matches(html);

            foreach (Match match in matches)
            {
                string textoAoRedor = match.Value;
                string codigoIbge = match.Groups[2].Value;

                // Checa se nas proximidades do código existe menção ao padrão nacional
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