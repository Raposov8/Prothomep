using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SGID.Data.ViewModel
{
    public class DadosCirurgia
    {
        [Key]
        public int Id { get; set; }
        public DateTime DataCriacao { get; set; }
        public string? Codigo { get; set; }
        public string NomePaciente { get; set; }
        public string NomeMedico { get; set; }
        public string NomeCliente { get; set; }
        public int Status { get; set; }
        public DateTime DataCirurgia { get; set; }
        public string? ProcedimentosExec { get; set; }
        public string? ObsIntercorrencia { get; set; }
        [ForeignKey("DadosCirurgiaId")]
        public ICollection<AnexosDadosCirurgia> Anexos { get; set; }
        [ForeignKey("AgendamentoId")]
        public int AgendamentoId { get; set; }
        public Agendamentos Agendamento { get; set; }
    }
}
