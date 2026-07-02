public class AppSettings
{
    public NddSettings Ndd { get; set; } = new();

    public ArquivosSettings Arquivos { get; set; } = new();
}

public class NddSettings
{
    public string Url { get; set; } = "";
}

public class ArquivosSettings
{
    public string PlanilhaCliente { get; set; } = "";

    public string PlanilhaResultado { get; set; } = "";
}