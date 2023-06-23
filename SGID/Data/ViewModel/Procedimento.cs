using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class Procedimento
    {
        [Key]
        public int Id { get; set; }
        public string Nome { get; set; }
        public decimal Valor { get; set; }
        public int Bloqueado { get; set; }
        public string MotivoBloqueio { get; set; }
        public string Empresa { get; set; }
    }
}
