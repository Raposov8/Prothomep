using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;
using System.ComponentModel.DataAnnotations;

namespace SGID.Pages.Account.RH
{
    [Authorize]
    public class EditarADMModel : PageModel
    {
        private readonly ApplicationDbContext _db;
        private readonly UserManager<UserInter> _userManager;
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        [BindProperty]
        public EditUserModel Editar { get; set; }
        public List<string> Linhas { get; set; } = new List<string>();
        public List<string> Gerente { get; set; } = new List<string> { "SIM", "NAO" };

        public EditarADMModel(ApplicationDbContext context, UserManager<UserInter> userManager
            , TOTVSDENUOContext DENUO, TOTVSINTERContext INTER)
        {
            _db = context;
            _userManager = userManager;
            Editar = new();
            ProtheusDenuo = DENUO;
            ProtheusInter = INTER;
        }
        public async Task<IActionResult> OnGet(int id)
        {
            if (id == 0)
            {
                return LocalRedirect("Error");
            }
            var lider = User.Identity.Name.Split("@")[0];
            var usuario = _db.TimeADMs.FirstOrDefault(x => x.Id == id);

            if (usuario == null)
            {
                return LocalRedirect("Error");
            }

            var user = await _userManager.FindByIdAsync(usuario.IdUsuario);
            if (user == null)
            {
                return LocalRedirect("Error");
            }

            var email = user.Email.Split("@")[1];
            Editar = new()
            {
                Id = user.Id,
                Email = user.Email,
                Porcentagem = usuario.Porcentagem,
            };

            return Page();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            try
            {
                var lider = User.Identity.Name.Split("@")[0];
                var usuario = _db.TimeADMs.FirstOrDefault(x => x.IdUsuario == Editar.Id);

                var user = await _userManager.FindByIdAsync(Editar.Id);
                if (user == null)
                {

                    ModelState.AddModelError("", "Usuário não encontrado");
                    return Page();
                }

                user.UserName = Editar.Email;
                user.Email = Editar.Email;

                user.NormalizedEmail = Editar.Email.ToUpper();
                user.NormalizedUserName = Editar.Email.ToUpper();

                usuario.Integrante = Editar.Email.Split("@")[0];

                usuario.Porcentagem = Editar.Porcentagem;


                _db.TimeADMs.Update(usuario);
                _db.SaveChanges();


                return Redirect("/account/rh/listartime");
            }
            catch (Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep, _db, "Edit", user);
            }
            ModelState.AddModelError("", "Something failed.");
            return Page();
        }

        public class EditUserModel
        {
            public string Id { get; set; }

            [Required(AllowEmptyStrings = false)]
            [Display(Name = "E-mail")]
            [EmailAddress]
            public string Email { get; set; }

            [Display(Name = "% Sobre Faturamento")]
            public double Porcentagem { get; set; }

        }
    }
}
