using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class PatrimonioAgendamento
    {
        [Key]
        public int Id { get; set; }
        [ForeignKey("AgendamentoId")]
        public int AgendamentoId { get; set; }
        public Agendamentos Agendamento { get; set; }
        public string Patrimonio { get; set; }
        public string? Codigo { get; set; }
    }
}
