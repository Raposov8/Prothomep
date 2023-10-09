using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class TimeADM
    {
        [Key]
        public int Id { get; set; }
        public DateTime DataCriacao { get; set; }
        public bool Status { get; set; }
        public string Integrante { get; set; }
        public double Porcentagem { get; set; }
        public string IdUsuario { get; set; }
        public string? Color { get; set; }
        public double Teto { get; set; }
        public double Salario { get; set; }
    }
}
