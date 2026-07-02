using HtmlAgilityPack;
using MunicipiosHomologados.ConsoleApp.Util;

namespace MunicipiosHomologados.ConsoleApp.Services
{
    public class NddService
    {
        private const string UrlNdd = "https://documentacao-nfse.e-datacenter.nddigital.com.br/fiscal-documentacao/docs/ndd-nfse/municipios-nfse/";
        private static readonly HttpClient _httpClient = new HttpClient();
        public async Task<Dictionary<string, bool>> ObterMunicipiosHomologadosAsync()
        {
            Logger.Info("🌐 Consultando site da NDD...");
            var resultado = await ObterMunicipiosHomologadosNddAsync(UrlNdd);
            Logger.Sucesso($"Mapeamento da NDD concluído! {resultado.Count} municípios carregados.");
            return resultado;
        }
        private static async Task<Dictionary<string, bool>> ObterMunicipiosHomologadosNddAsync(string url)
        {
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) C# Core");

            string html = await _httpClient.GetStringAsync(url);
            if (!html.Contains("Municípios Homologados", StringComparison.OrdinalIgnoreCase))
                throw new Exception("A página da NDD não retornou o conteúdo esperado.");


            var mapaMunicipios = new Dictionary<string, bool>();

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

                Console.WriteLine($"Município: {colunas[0].InnerText.Trim()}");
                Console.WriteLine($"UF: {colunas[1].InnerText.Trim()}");
                Console.WriteLine($"IBGE: {colunas[2].InnerText.Trim()}");
                Console.WriteLine($"Padrão: {colunas[3].InnerText.Trim()}");

                string codigoIbge = HtmlEntity.DeEntitize(colunas[2].InnerText).Trim();
                string padrao = HtmlEntity.DeEntitize(colunas[3].InnerText).Trim();

                bool ehNacional = padrao.Equals(
                    "NFSeNacional",
                    StringComparison.OrdinalIgnoreCase);

                mapaMunicipios[codigoIbge] = ehNacional;
            }

            return mapaMunicipios;
        }
    }
}