using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using MunicipiosHomologados.ConsoleApp.Models;
using MunicipiosHomologados.ConsoleApp.Util;
using Microsoft.Extensions.Configuration;
using MunicipiosHomologados.ConsoleApp.Configuration;

namespace MunicipiosHomologados.ConsoleApp.Services
{
    public class ExcelService
    {
        // =========================================================================
        // 1. PROCESSA A PLANILHA DO CLIENTE (CRUZAMENTO DE DADOS)
        // =========================================================================
        public void ProcessarPlanilhaClientes(Dictionary<string, MunicipioIbge> dicionarioIbge, List<MunicipioNdd> municipiosNdd)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Uso Pessoal");

            var settings = new AppSettings();
            AppConfig.Configuration.Bind(settings);

            string caminhoOrigem = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, settings.Arquivos.PlanilhaCliente);
            string caminhoDestino = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, settings.Arquivos.PlanilhaResultado);

            // [ADICIONE ESTA LINHA] -> Protege a gravação do resultado final
            GarantirArquivoDisponivel(caminhoDestino);

            if (!File.Exists(caminhoOrigem))
            {
                throw new FileNotFoundException($"A planilha do cliente não foi encontrada em: {caminhoOrigem}");
            }

            Logger.Info("📊 Cruzando dados e gerando planilha de resultados...");

            using (var package = new ExcelPackage(new FileInfo(caminhoOrigem)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int totalLinhas = worksheet.Dimension.End.Row;

                // Cria os novos cabeçalhos nas Colunas C, D e E
                worksheet.Cells[1, 3].Value = "Código IBGE";
                worksheet.Cells[1, 4].Value = "Homologado NDD";
                worksheet.Cells[1, 5].Value = "Nacional";
                worksheet.Cells[1, 6].Value = "Padrão";

                // Estilização do cabeçalho completo (Colunas A até F)
                using (var range = worksheet.Cells[1, 1, 1, 6])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Name = "Arial";
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Navy);
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                for (int linha = 2; linha <= totalLinhas; linha++)
                {
                    string municipioCliente = worksheet.Cells[linha, 1].Value?.ToString()?.ToUpper().Trim() ?? "";
                    string ufCliente = worksheet.Cells[linha, 2].Value?.ToString()?.ToUpper().Trim() ?? "";

                    // Chave combinada idêntica à do IBGE (Ex: "SÃO PAULO - SP")
                    string chaveBusca = $"{municipioCliente} - {ufCliente}";

                    // 1. Busca o Código IBGE em memória
                    if (dicionarioIbge.TryGetValue(chaveBusca, out var dadosIbge))
                    {
                        worksheet.Cells[linha, 3].Value = dadosIbge.CodigoFormatado;

                        // 2. Busca na lista da NDD priorizando a linha que seja "NFSeNacional"
                        var homologado = municipiosNdd
                            .Where(x => x.CodigoIbge == dadosIbge.CodigoFormatado)
                            .OrderByDescending(x => x.Padrao.Equals("NFSeNacional", StringComparison.OrdinalIgnoreCase))
                            .FirstOrDefault();

                        if (homologado != null)
                        {
                            worksheet.Cells[linha, 4].Value = "SIM";
                            worksheet.Cells[linha, 6].Value = homologado.Padrao;

                            // Verifica se o padrão retornado é o nacional para preencher a nova coluna
                            bool ehNacional = homologado.Padrao.Equals("NFSeNacional", StringComparison.OrdinalIgnoreCase);
                            worksheet.Cells[linha, 5].Value = ehNacional ? "SIM" : "NÃO";

                            // Pinta a linha de Verde Claro se estiver homologado
                            worksheet.Cells[linha, 1, linha, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[linha, 1, linha, 6].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(220, 245, 220));
                        }
                        else
                        {
                            worksheet.Cells[linha, 4].Value = "NÃO";
                            worksheet.Cells[linha, 5].Value = "NÃO";
                            worksheet.Cells[linha, 6].Value = "N/A";

                            // Pinta de Vermelho bem claro se não estiver homologado
                            worksheet.Cells[linha, 1, linha, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[linha, 1, linha, 6].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 225, 225));
                        }
                    }
                    else
                    {
                        // Se não achar o município no IBGE
                        worksheet.Cells[linha, 3].Value = "Não Encontrado no IBGE";
                        worksheet.Cells[linha, 4].Value = "N/A";
                        worksheet.Cells[linha, 5].Value = "N/A";
                    }

                    // Aplica efeito zebrado nas linhas que não foram pintadas por status
                    if (worksheet.Cells[linha, 4].Value?.ToString() == "N/A" && linha % 2 == 0)
                    {
                        worksheet.Cells[linha, 1, linha, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[linha, 1, linha, 5].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 245, 249));
                    }
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                package.SaveAs(new FileInfo(caminhoDestino));
            }
        }

        // =========================================================================
        // 2. GERA A PLANILHA ESPELHO COM TODOS OS HOMOLOGADOS DA NDD (BACKUP)
        // =========================================================================
        public void GerarPlanilhaBaseNdd(List<MunicipioNdd> municipios, string caminhoDestino)
        {
            // [ADICIONE ESTA LINHA] -> Valida e espera o usuário fechar se estiver aberto
            GarantirArquivoDisponivel(caminhoDestino);
            ExcelPackage.License.SetNonCommercialPersonal("Uso Pessoal");

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Homologados NDD");

                // Cabeçalhos
                worksheet.Cells[1, 1].Value = "Município";
                worksheet.Cells[1, 2].Value = "UF";
                worksheet.Cells[1, 3].Value = "Código IBGE";
                worksheet.Cells[1, 4].Value = "Padrão Tecnológico";

                // Estilizar Cabeçalho (Navy / Branco / Arial Bold)
                using (var range = worksheet.Cells[1, 1, 1, 4])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Name = "Arial";
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Navy);
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                int linha = 2;
                foreach (var m in municipios)
                {
                    worksheet.Cells[linha, 1].Value = m.Nome;
                    worksheet.Cells[linha, 2].Value = m.Uf;
                    worksheet.Cells[linha, 3].Value = m.CodigoIbge;
                    worksheet.Cells[linha, 4].Value = m.Padrao;

                    // Efeito zebrado nas linhas pares
                    if (linha % 2 == 0)
                    {
                        using (var rangeLinha = worksheet.Cells[linha, 1, linha, 4])
                        {
                            rangeLinha.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            rangeLinha.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 245, 249));
                        }
                    }
                    linha++;
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                package.SaveAs(new FileInfo(caminhoDestino));
            }
        }

        private void GarantirArquivoDisponivel(string caminhoArquivo)
        {
            var arquivoInfo = new FileInfo(caminhoArquivo);

            // Se o arquivo nem existe ainda, não tem o que validar
            if (!arquivoInfo.Exists) return;

            bool arquivoTravado = true;

            while (arquivoTravado)
            {
                try
                {
                    // Tenta abrir o arquivo com acesso exclusivo de escrita.
                    // Se falhar, vai direto para o catch.
                    using (FileStream stream = arquivoInfo.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    {
                        arquivoTravado = false; // Conseguiu abrir? Então não está travado!
                    }
                }
                catch (IOException)
                {
                    Logger.Aviso($"⚠️  O arquivo '{arquivoInfo.Name}' está aberto no Excel.");
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("👉 Por favor, feche a planilha e pressione [ENTER] para tentar salvar novamente...");
                    Console.ResetColor();

                    Console.ReadLine(); // Trava a execução até o usuário dar Enter
                }
            }
        }
    }
}