using System.ComponentModel.DataAnnotations;
using System.Data;

namespace SGID.Data.ViewModel
{
    public class Dispositivo
    {
        [Key]
        public int Id { get; set; }
        public DateTime DataCadastro { get; set; }
        public DateTime? DataAlteracao { get; set; }
        public string Nome { get; set; }
        public string Modelo { get; set; }
        public string? Imei { get; set; }
        public string TipoDispositivo { get; set; }
        public double Valor { get; set; }

        public string UsuarioCriacao { get; set; }
        public string? UsuarioAlteracao { get; set; }
    }
}
