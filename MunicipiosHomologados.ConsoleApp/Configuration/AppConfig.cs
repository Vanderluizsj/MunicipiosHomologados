using Microsoft.Extensions.Configuration;

namespace MunicipiosHomologados.ConsoleApp.Configuration;

public static class AppConfig
{
    public static IConfigurationRoot Configuration { get; } =
        new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false)
            .Build();
}