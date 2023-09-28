using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class TimeDental
    {

        [Key]
        public int Id { get; set; }
        public DateTime DataCriacao { get; set; }
        public bool Status { get; set; }
        public string Integrante { get; set; }
        public string Lider { get; set; }
        public double Meta { get; set; }
        public double Porcentagem { get; set; }
        public double PorcentagemSeg { get; set; }
        public string IdUsuario { get; set; }
        public string? Color { get; set; }
    }
}
