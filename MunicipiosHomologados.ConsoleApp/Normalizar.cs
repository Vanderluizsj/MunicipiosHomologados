using System.Globalization;
using System.Text;

namespace MunicipiosHomologados.ConsoleApp;

public static class Normaliza
{
    public static string Normalizar(string texto)
    {
        // Remove acentos e caracteres especiais
        string textoNormalizado = texto.Normalize(NormalizationForm.FormD);
        StringBuilder sb = new StringBuilder();
        foreach (char c in textoNormalizado)
        {
            UnicodeCategory uc = CharUnicodeInfo.GetUnicodeCategory(c);
            if (uc != UnicodeCategory.NonSpacingMark)
            {
                sb.Append(c);
            }
        }
        return sb.ToString().Normalize(NormalizationForm.FormC).ToLower();
    }
}