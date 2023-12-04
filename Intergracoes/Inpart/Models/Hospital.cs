using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.Inpart.Models
{
    public class Hospital
    {
        public int? idHospital { get; set; }
        public string? razaoSocial { get; set; }
        public string? cgc { get; set; }
        public string? inscricaoEstadual { get; set; }
        public string? endereco { get; set; }
        public string? cidade { get; set; }
        public string? bairro { get; set; }
        public string? uf { get; set; }
        public string? cep { get; set; }
        public string? ddd { get; set; }
        public string? fone { get; set; }
    }
}
