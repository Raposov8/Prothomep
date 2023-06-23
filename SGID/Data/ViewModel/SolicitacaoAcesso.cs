using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class SolicitacaoAcesso
    {
        [Key]
        public int Id { get; set; }
        public DateTime DataCriacao { get; set; }
        public DateTime? DataAlteracao { get; set; }
        public DateTime DataEvento { get; set; }
        public string Empresa { get; set; }
        public string Nome { get; set; } 
        public string Tipo { get; set; }
        public string Email { get; set; }
        public string Cargo { get; set; }
        public string? Ramal { get; set; }
        public string? Obs { get; set; }
        public bool Protheus { get; set; }
        public bool IsRamal { get; set; }
        public bool Maquina { get; set; }
        public bool Celular { get; set; }
        public bool Impressora { get; set; }
        public string Usuario { get; set; }
    }
}
