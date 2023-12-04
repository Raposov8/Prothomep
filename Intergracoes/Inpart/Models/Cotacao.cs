using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.Inpart.Models
{
    public class Cotacao
    {
        
        public int? idCotacao { get; set; }
        public int? idSeguradora { get; set; }
        public int? idSucursal { get; set; }
        public int? idCliente { get; set; }
        public int? idHospital { get; set; }
        public int? nuGuia { get; set; }
        public int? idMedico { get; set; }
        public DateTime? dtCotacao { get; set; }
        public int? cdStatus { get; set; }
        public DateTime? dtCirurgia { get; set; }
        public string? nmHoraCirurgia { get; set; }
        public int? codAssociado { get; set; }
        public int? codCirurgia { get; set; }
        public int? codInternacao { get; set; }
        public int? codPaciente { get; set; }
        public List<ItemCotacao> cotacaoItemList { get; set; }
        public Hospital hospital { get; set; }
        public Medico medico { get; set; }
        
   
    }
}
