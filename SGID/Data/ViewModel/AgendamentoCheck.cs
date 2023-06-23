using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SGID.Data.ViewModel
{
    public class AgendamentoCheck
    {
        [Key]
        public int Id { get; set; }
        [ForeignKey("AgendamentoId")]
        public int AgendamentoId { get; set; }
        public Agendamentos Agendamento { get; set; }
        public string Codigo { get; set; }
        public string Descricao { get; set; }
        public double Quantidade { get; set; }
        public bool Entregue { get; set; }

    }
}
