using Microsoft.AspNetCore.Identity;

namespace SGID.Data.ViewModel
{
    public class UserInter : IdentityUser
    {
        public string? NomeVendedor { get; set; }
        public string? UsuarioCriacao { get; set; }
        public DateTime? CriacaoDate { get; set; }
        public string? UsuarioAlterar { get; set; }
        public DateTime? AlterarDate { get; set; }
    }
}
