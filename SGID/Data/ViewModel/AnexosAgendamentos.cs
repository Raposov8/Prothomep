using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SGID.Data.ViewModel
{
    public class AnexosAgendamentos
    {
        [Key]
        public int Id { get; set; }
        [ForeignKey("AgendamentoId")]
        public int AgendamentoId { get; set; }
        public Agendamentos Agendamento { get; set; }
        public string NumeroAnexo { get; set; }
        public string AnexoCam { get; set; }
    }
}
