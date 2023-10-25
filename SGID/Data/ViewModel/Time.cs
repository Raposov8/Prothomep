using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class Time
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
        public double PorcentagemGenProd { get; set; }
        public string IdUsuario { get; set; }
        public string GerenProd { get; set; }
        public string? Color { get; set; }
        public string? TipoFaturamento { get; set; }
        public string? TipoVendedor { get; set; }
        public double Teto { get; set; }
        public double Salario { get; set; }
        public bool PJ { get; set; }
    }
}
