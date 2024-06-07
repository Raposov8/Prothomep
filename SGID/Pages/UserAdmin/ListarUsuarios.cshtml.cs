using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;

namespace SGID.Pages.UserAdmin
{
    [Authorize(Roles = "Admin")]
    public class ListarUsuariosModel : PageModel
    {
        private readonly UserManager<UserInter> _userManager;

        public List<UserInter> Users { get; set; }
        public ListarUsuariosModel(UserManager<UserInter> manager)
        {
            _userManager = manager;
            Users = new();
        }
        public void OnGet(int id)
        {

            Users = _userManager.Users.OrderBy(x=>x.UserName).ToList();
        }
    }
}
