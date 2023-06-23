using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection.Metadata;

namespace SGID.Data.ViewModel
{
    public class ObsAgendamento
    {
        [Key]
        public int Id { get; set; }
        [ForeignKey("AgendamentoId")]
        public int AgendamentoId { get; set; }
        public Agendamentos Agendamento { get; set; }
        public DateTime DataCriacao { get; set; }
        public string User { get; set; }
        public string Obs { get; set; }
    }
}
