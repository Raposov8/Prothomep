using Microsoft.AspNetCore.Identity;

namespace SGID.Data.ViewModel
{
    public class UserInter : IdentityUser
    {
        public string? NomeVendedor { get; set; }
    }
}
