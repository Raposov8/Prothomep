using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class Ocorrencia
    {
        [Key]
        public int Id { get; set; }
        public DateTime DataCriacao { get; set; }
        public string? Empresa { get; set; }
        public string? Cliente { get; set; }
        public string? Hospital { get; set; }
        public string? Medico { get; set; }
        public string? Paciente { get; set; }
        public string? Agendamento { get; set; }
        public DateTime? Cirurgia { get; set; }
        public string? Patrimonio { get; set; }
        public string? DescPatri { get; set; }
        public string? Produto { get; set; }
        public string? Descricao { get; set; }
        public int Quantidade { get; set; }
        public DateTime? DataOcorrencia { get; set; }
        public string? Problema { get; set; }
        public string? Acao { get; set; }
        public string? Procedente { get; set; }
        public string? Cobrado { get; set; }
        public string? Reposto { get; set; }
        public string? Vendedor { get; set; }
        public string? Motorista { get; set; }
        public string? Obs { get; set; }

        public string? UsuarioCriacao { get; set; }
    }
}
