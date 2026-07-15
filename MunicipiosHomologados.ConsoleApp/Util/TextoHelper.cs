using System.Globalization;
using System.Text;

namespace MunicipiosHomologados.ConsoleApp.Util;


public static class TextoHelper
{
    public static string Normalizar(string? texto)
    {
        if (string.IsNullOrWhiteSpace(texto))
            return string.Empty;

        texto = texto.Trim().ToUpperInvariant();

        texto = texto.Normalize(NormalizationForm.FormD);

        var sb = new StringBuilder();

        foreach (char c in texto)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                sb.Append(c);
        }

        texto = sb.ToString().Normalize(NormalizationForm.FormC);

        texto = texto.Replace("'", "")
                     .Replace("’", "")
                     //.Replace("-", " ")
                     .Replace("/", " ")
                     .Replace(".", " ")
                     .Replace(",", " ");

        while (texto.Contains("  "))
            texto = texto.Replace("  ", " ");

        return texto;
    }
}