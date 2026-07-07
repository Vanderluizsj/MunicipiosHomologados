using HtmlAgilityPack;
using Microsoft.Extensions.Configuration;
using MunicipiosHomologados.ConsoleApp.Configuration;
using MunicipiosHomologados.ConsoleApp.Models;
using MunicipiosHomologados.ConsoleApp.Util;

namespace MunicipiosHomologados.ConsoleApp.Services
{
    public class NddService
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        public async Task<List<MunicipioNdd>> ObterMunicipiosHomologadosAsync()
        {
            var settings = new AppSettings();
            AppConfig.Configuration.Bind(settings);

            Logger.Info("🌐 Consultando site da NDD...");
            var resultado = await ObterMunicipiosHomologadosNddAsync(settings.Ndd.Url);
            Logger.Sucesso($"Mapeamento da NDD concluído! {resultado.Count} municípios localizados.");

            return resultado;
        }

        private static async Task<List<MunicipioNdd>> ObterMunicipiosHomologadosNddAsync(string url)
        {
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) C# Core");

            string html = await _httpClient.GetStringAsync(url);
            if (!html.Contains("Municípios Homologados", StringComparison.OrdinalIgnoreCase))
                throw new Exception("A página da NDD não retornou o conteúdo esperado.");

            var dicionarioTemporario = new Dictionary<string, MunicipioNdd>();

            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            var linhas = doc.DocumentNode.SelectNodes("//table/tbody/tr");
            if (linhas == null)
                throw new Exception("Nenhuma linha da tabela foi encontrada na página.");

            foreach (var linha in linhas)
            {
                var colunas = linha.SelectNodes("td");
                if (colunas == null || colunas.Count < 4)
                    continue;

                string codigoIbge = HtmlEntity.DeEntitize(colunas[2].InnerText).Trim();
                string padrao = HtmlEntity.DeEntitize(colunas[3].InnerText).Trim();

                var novoMunicipio = new MunicipioNdd
                {
                    Nome = HtmlEntity.DeEntitize(colunas[0].InnerText).Trim(),
                    Uf = HtmlEntity.DeEntitize(colunas[1].InnerText).Trim(),
                    CodigoIbge = codigoIbge,
                    Padrao = padrao
                };

                if (dicionarioTemporario.TryGetValue(codigoIbge, out var municipioExistente))
                {
                    // Atribui pesos para definir a prioridade (quanto maior o peso, mais prioritário)
                    int pesoExistente = ObterPesoPadrao(municipioExistente.Padrao);
                    int pesoNovo = ObterPesoPadrao(novoMunicipio.Padrao);

                    // Se o novo padrão tiver uma prioridade maior do que o que já estava lá, substitui!
                    if (pesoNovo > pesoExistente)
                    {
                        dicionarioTemporario[codigoIbge] = novoMunicipio;
                    }
                }
                else
                {
                    dicionarioTemporario.Add(codigoIbge, novoMunicipio);
                }
            }

            return dicionarioTemporario.Values.ToList();
        }

        // Função auxiliar interna para calcular a prioridade dos padrões
        private static int ObterPesoPadrao(string padrao)
        {
            if (padrao.Equals("NFSeNacional", StringComparison.OrdinalIgnoreCase))
                return 3; // Prioridade Máxima

            if (padrao.EndsWith("_REFORMA", StringComparison.OrdinalIgnoreCase))
                return 2; // Segunda maior prioridade

            return 1; // Padrão comum/antigo
        }
    }
}