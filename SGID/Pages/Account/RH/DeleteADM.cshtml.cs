using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;

namespace SGID.Pages.Account.RH
{
    [Authorize]
    public class DeleteADMModel : PageModel
    {
        private readonly UserManager<UserInter> _userManager;
        private readonly ApplicationDbContext _db;

        [BindProperty]
        public Delete Input { get; set; }

        public DeleteADMModel(UserManager<UserInter> userManager, ApplicationDbContext db)
        {
            _userManager = userManager;
            _db = db;
            Input = new();
        }

        public async Task<ActionResult> OnGet(int id)
        {
            try
            {
                var lider = User.Identity.Name.Split("@")[0];
                var Integrante = _db.TimeADMs.FirstOrDefault(x => x.Id == id);

                var user = await _userManager.FindByIdAsync(Integrante.IdUsuario);

                Input = new Delete
                {
                    Id = Integrante.Id,
                    UserName = user.UserName,
                    IdUsuario = user.Id
                };


                return Page();
            }
            catch (Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep, _db, "DeleteTime, Get", user);
                return Redirect("/Account/ListarTime");
            }

        }

        public async Task<IActionResult> OnPostAsync(int id)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    var Integrante = _db.TimeADMs.FirstOrDefault(x => x.Id == id);

                    var user = await _userManager.FindByIdAsync(Integrante.IdUsuario);

                    Integrante.Status = false;
                    Integrante.Desativar = Input.DataDesativacao;

                    _db.TimeADMs.Update(Integrante);
                    _db.SaveChanges();

                    return Redirect("/Account/RH/listartime");
                }
            }
            catch (Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep, _db, "DeleteTime, Post", user);
            }
            return Page();
        }

        public class Delete
        {
            public int Id { get; set; }
            public string UserName { get; set; }
            public DateTime DataDesativacao { get; set; }
            public string IdUsuario { get; set; }
        }
    }
}
