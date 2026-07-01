using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MunicipiosHomologados.ConsoleApp
{
    public class ResultadoValidacao
    {
        public int LinhaExcel { get; set; }
        public string Status { get; set; } = "";
        public string Nacional { get; set; } = string.Empty;
    }
}