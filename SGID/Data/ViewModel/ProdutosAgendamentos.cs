using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SGID.Data.ViewModel
{
    public class ProdutosAgendamentos
    {
        [Key]
        public int Id { get; set; }
        [ForeignKey("AgendamentoId")]
        public int AgendamentoId { get; set; }
        public Agendamentos Agendamento { get; set; }
        public string CodigoProduto { get; set; }
        public string CodigoTabela { get; set; }
        public double Quantidade { get; set; }
        public double ValorUnitario { get; set; }
        public double ValorTotal { get; set; }
    }
}
