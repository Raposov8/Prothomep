using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SGID.Data.ViewModel
{
    public class DadosCirugiasProdutos
    {
        [Key]
        public int Id { get; set; }
        [ForeignKey("DadosCirurgiaId")]
        public int DadosCirurgiaId { get; set; }
        public DadosCirurgia DadosCirurgia { get; set; }
        public string Produto { get; set; }
        public double Quantidade { get; set; }
    }
}
