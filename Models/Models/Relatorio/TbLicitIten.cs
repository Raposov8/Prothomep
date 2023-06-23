using System;
using System.Collections.Generic;

namespace SGID.Models.Relatorio
{
    public partial class TbLicitIten
    {
        public string? Numproc { get; set; }
        public string? Empenho { get; set; }
        public int? Grupo { get; set; }
        public string? Item { get; set; }
        public string? Descitem { get; set; }
        public int? Qtde { get; set; }
        public double? Vlunit { get; set; }
        public string? Codcli { get; set; }
        public string? Lojacli { get; set; }
    }
}
