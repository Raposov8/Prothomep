using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SGID.Data.ViewModel
{
    public class FormularioAvulsoXProdutos
    {
        [Key]
        public int Id { get; set; }
        [ForeignKey("AgendamentoId")]
        public int FormularioId { get; set; }
        public FormularioAvulso Formulario { get; set; }
        public string Produto { get; set; }
        public string Descricao { get; set; }
        public double Quantidade { get; set; }
        public string Lote { get;set;}

    }
}
