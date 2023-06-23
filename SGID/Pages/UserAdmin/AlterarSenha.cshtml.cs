using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;

namespace SGID.Pages.UserAdmin
{
    [Authorize(Roles = "Admin")]
    public class AlterarSenhaModel : PageModel
    {
        private readonly UserManager<UserInter> _userManager;
        private readonly ApplicationDbContext _db;

        [BindProperty]
        public AlterarSenha Senha { get; set; }

        public AlterarSenhaModel(UserManager<UserInter> manager,ApplicationDbContext context)
        {
            _userManager = manager;
            _db = context;
            Senha = new();
        }
        public async Task<IActionResult> OnGet(string id)
        {
            try
            {
                var user = await _userManager.FindByIdAsync(id);
                var usuario = _db.Users.Where(l => l.Id == user.Id).FirstOrDefault();

                Senha = new AlterarSenha 
                { 
                   Id = usuario.Id,
                   Email = usuario.Email,
                };

                return Page();
            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex,_db,"AlterarSenha,Get",user);
                return Redirect("/UserAdmin/ListarUsuarios/1");
            }
        }

        public async Task<IActionResult> OnPostAsync()
        {
            try
            {
                if (ModelState.IsValid)
                {
                    var user = await _userManager.FindByIdAsync(Senha.Id);

                    var result = await _userManager.RemovePasswordAsync(user);

                    result = await _userManager.AddPasswordAsync(user, Senha.Password);

                    if (!result.Succeeded)
                    {
                        ModelState.AddModelError("", result.Errors.First().Description);
                        return Page();
                    }

                    return Redirect("/UserAdmin/ListarUsuarios/1");
                }
            }
            catch (Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep, _db, "AlterarSenha,Post",user);
            }
            ModelState.AddModelError("", "Something failed.");
            return Page();
        }

        public class AlterarSenha
        {
            public string Id { get; set; }
            public string Password { get; set; }
            public string Email { get; set; }
        }
    }
}
