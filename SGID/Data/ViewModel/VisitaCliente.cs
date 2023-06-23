using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class VisitaCliente
    {
        [Key]
        public int Id { get; set; }
        public DateTime DataCriacao { get; set; }
        public string Medico { get; set; }
        public string? Local { get; set; }
        public string? Assunto { get; set; }
        public string? Endereco { get; set; }
        public string? Bairro { get; set; }
        public string Empresa { get; set; }
        public string Vendedor { get; set; }

    }
}
