using System.ComponentModel.DataAnnotations.Schema;

namespace SGID.Data.ViewModel
{
    public class UsuarioDispositivo
    {
        public int Id { get; set; }
        public string NomeUsuario { get; set; }
        public int DispositivoId { get; set; }
        public Dispositivo Dispositivo { get; set; }
        public bool Ativo { get; set; }
    }
}
