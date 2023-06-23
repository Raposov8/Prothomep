using System;
using System.Collections.Generic;

namespace SGID.Models.Relatorio
{
    public partial class Apagar
    {
        public string? Empresa { get; set; }
        public string? Anomes { get; set; }
        public string? Natureza { get; set; }
        public string? Descnat { get; set; }
        public string? Grpnat { get; set; }
        public string? Descgrupo { get; set; }
        public string? Codforn { get; set; }
        public string? Lojaforn { get; set; }
        public string? Nomforn { get; set; }
        public double? Vlpagar { get; set; }
        public double? Vlpago { get; set; }
    }
}
