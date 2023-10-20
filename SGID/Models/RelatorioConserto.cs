using System;

namespace SGID.Models
{
    public class RelatorioConserto
    {
        public int Dias { get; set; }
        public string Filial { get; set; }
        public string Nf { get; set; }
        public string Serie { get; set; }
        public DateTime? Emissao { get; set; }
        public string CodCli { get; set; }
        public string Loja { get; set; }
        public string Cliente { get; set; }
        public string CF { get; set; }
    }
}
