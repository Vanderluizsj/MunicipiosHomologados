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
            ExcelPackage.License.SetNonCommercialPersonal("<Luiz>");
            try
            {
                Logger.Info("Iniciando processamento de alta performance estruturado...");

                var nddService = new NddService();
                var ibgeService = new IbgeService();
                var excelService = new ExcelService();

                // 1. NDD SERVICE: Baixa todos os dados da web para a memória
                var dadosNdd = await nddService.ObterMunicipiosHomologadosAsync();

                // [NOVO] 1.5. EXCEL SERVICE: Exporta a cópia idêntica do site para um arquivo Excel local de backup
                string caminhoBackupNdd = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Base_Atualizada_NDD.xlsx");
                excelService.GerarPlanilhaBaseNdd(dadosNdd, caminhoBackupNdd);
                Logger.Sucesso($"Planilha de backup da NDD criada em: {caminhoBackupNdd}");

                // 2. IBGE SERVICE: Carrega o dicionário com os códigos do IBGE
                var dicionarioIbge = ibgeService.CarregarMunicipios();

                // 3. EXCEL SERVICE: Processa o Procv e gera o resultado do cliente
                // (Ajustado para passar a lista 'dadosNdd' para dentro do processamento)
                excelService.ProcessarPlanilhaClientes(dicionarioIbge, dadosNdd);

                Logger.Sucesso("Processamento concluído com sucesso total!");
            }
            catch (Exception ex)
            {
                Logger.Info($"\n❌ Erro crítico na execução da orquestração principal: {ex.Message}");
                Logger.FatalErro(ex.ToString());
            }
        }
    }
}