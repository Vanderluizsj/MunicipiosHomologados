using MunicipiosHomologados.ConsoleApp.Services;
using MunicipiosHomologados.ConsoleApp.Util;
using OfficeOpenXml;

namespace MunicipiosHomologados.ConsoleApp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // Configuração da Licença (Garante que se aplique no escopo do serviço também)
            ExcelPackage.License.SetNonCommercialPersonal("<Vander>");
            try
            {
                Logger.Info("🚀 Iniciando processamento de alta performance estruturado...");

                // 1. CARREGA O SITE DA NDD EM SEGUNDO PLANO
                var nddService = new NddService();
                var municipiosHomologadosNdd = await nddService.ObterMunicipiosHomologadosAsync();

                // 2. CARREGA A BASE DO IBGE EMBUTIDA DENTRO DO EXE (PROCV EM MEMÓRIA)
                var ibgeService = new IbgeService(); // Certifique-se de que o método dele retorne Dictionary<string, string>
                var dicionarioIbge = ibgeService.CarregarMunicipios();

                // 3. PROCESSA, FORMATA E EXPORTA A PLANILHA FINAL
                var excelService = new ExcelService();
                excelService.ProcessarPlanilhaClientes(dicionarioIbge, municipiosHomologadosNdd);
                
                Logger.Info("\n🎉 Processamento concluído com sucesso total!");
            }
            catch (Exception ex)
            {
                Logger.Info($"\n❌ Erro crítico na execução da orquestração principal: {ex.Message}");
            }
        }
    }
}