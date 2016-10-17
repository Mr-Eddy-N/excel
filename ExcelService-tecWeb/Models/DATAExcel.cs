using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelService_tecWeb.Models
{
    public class DATAExcel
    {
        public List<string> Titulos { get; set;}
        public List<List<string>> Contenido { get; set;}
        public string Nombre { get; set; }
        public string Path { get; set; }
    }
}