using System;
using System.IO;
using System.Linq;
using System.Threading;
using ClosedXML.Excel;
using MunicipiosHomologados.ConsoleApp;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI; // Alterado para Edge

class Program
{
    static void Main(string[] args)
    {
        string caminhoExcel = @"C:\Source\MunicipiosHomologados\Cliente.xlsx"; // Caminho do arquivo Excel
        var resultados = new List<ResultadoValidacao>();
        var encontrado = false;
        EdgeOptions options = new EdgeOptions();

        using (IWebDriver driver = new EdgeDriver(options))
        {
            // Abre o site da NDD Space
            driver.Navigate().GoToUrl("https://documentacao-nfse.e-datacenter.nddigital.com.br/fiscal-documentacao/docs/ndd-nfse/municipios-nfse/");

            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            wait.Until(d =>
                d.FindElement(By.XPath("//input[@placeholder='Digite para buscar...']")));

            // Resto do código permanece exatamente igual...
            using (var workbook = new XLWorkbook(caminhoExcel))
            {
                var worksheet = workbook.Worksheet(1);
                var linhas = worksheet.RangeUsed().RowsUsed().Skip(1);


                foreach (var linha in linhas)
                {
                    string municipioProcurado = linha.Cell(1).GetString().Trim();
                    string ufProcurada = linha.Cell(2).GetString().Trim().ToUpper();

                    if (string.IsNullOrEmpty(municipioProcurado)) continue;

                    Console.WriteLine($"Validando: {municipioProcurado} - {ufProcurada}...");

                    try
                    {
                        IWebElement campoBusca = driver.FindElement(
     By.XPath("//input[@placeholder='Digite para buscar...']")
 );

                        campoBusca.Click();

                        campoBusca.SendKeys(Keys.Control + "a");
                        campoBusca.SendKeys(Keys.Delete);

                        campoBusca.SendKeys(municipioProcurado);

                        // pequena espera para o Angular atualizar
                        Thread.Sleep(300);

                        resultados.Add(new ResultadoValidacao
                        {
                            LinhaExcel = linha.RowNumber(),
                            Status = encontrado ? "Não Homologado" : "Homologado"
                        });

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erro ao processar {municipioProcurado}: {ex.Message}");
                        resultados.Add(new ResultadoValidacao
                        {
                            LinhaExcel = linha.RowNumber(),
                            Status = "Erro na Validação"
                        });
                    }
                }
                foreach (var resultado in resultados)
                {
                    worksheet.Cell(resultado.LinhaExcel, 3).Value = resultado.Status;
                }

                workbook.Save();

            }

            driver.Quit();
        }

        Console.WriteLine("\nProcesso concluído com sucesso! Planilha salva.");
    }
}