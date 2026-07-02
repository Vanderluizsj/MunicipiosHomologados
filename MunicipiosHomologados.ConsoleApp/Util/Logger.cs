namespace MunicipiosHomologados.ConsoleApp.Util
{
    public static class Logger
    {
        // 1. Define o caminho da subpasta "Logs" e o arquivo dentro dela
        private static readonly string PastaLog = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
        private static readonly string CaminhoLog = Path.Combine(PastaLog, "execucao.log");

        public static void Info(string mensagem)
        {
            Logar(mensagem, ConsoleColor.White, "INFO");
        }

        public static void Sucesso(string mensagem)
        {
            Logar(mensagem, ConsoleColor.Green, "SUCESSO");
        }

        public static void Aviso(string mensagem)
        {
            Logar(mensagem, ConsoleColor.Yellow, "AVISO");
        }

        public static void Erro(string mensagem)
        {
            Logar(mensagem, ConsoleColor.Red, "ERRO");
        }

        private static void Logar(string mensagem, ConsoleColor cor, string nivel)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string linhaTextoLog = $"[{timestamp}] [{nivel}] {mensagem}";

            try
            {
                // 2. VERIFICAÇÃO CRUCIAL: Se a pasta "Logs" não existir, o C# cria ela na hora
                if (!Directory.Exists(PastaLog))
                {
                    Directory.CreateDirectory(PastaLog);
                }

                // 3. Grava o texto de forma segura dentro da subpasta
                File.AppendAllText(CaminhoLog, linhaTextoLog + Environment.NewLine);
            }
            catch
            {
                // Protege o fluxo caso o arquivo de log esteja aberto ou travado
            }

            // 4. Exibe no Console nativo com a cor configurada
            Console.ForegroundColor = cor;
            Console.WriteLine(linhaTextoLog);
            Console.ResetColor();
        }
    }
}