using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;
using SGID.Data;
using SGID.Data.Models;

namespace SGID.Pages.Account.RH
{
    public class DeleteDentalModel : PageModel
    {
        private readonly UserManager<UserInter> _userManager;
        private readonly ApplicationDbContext _db;

        public Delete Input { get; set; }

        public DeleteDentalModel(UserManager<UserInter> userManager, ApplicationDbContext db)
        {
            _userManager = userManager;
            _db = db;
            Input = new();
        }

        public async Task<ActionResult> OnGet(int id)
        {
            try
            {
                var Integrante = _db.TimeDentals.FirstOrDefault(x => x.Id == id);

                //var user = await _userManager.FindByIdAsync(Integrante.IdUsuario);

                Input = new Delete
                {
                    Id = Integrante.Id,
                    UserName = Integrante.Integrante,
                    //IdUsuario = user.Id
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
                    var Integrante = _db.Times.FirstOrDefault(x => x.Id == id);

                    //var user = await _userManager.FindByIdAsync(Integrante.IdUsuario);

                    Integrante.Status = false;
                    Integrante.Desativar = DateTime.Now;

                    _db.Times.Update(Integrante);
                    _db.SaveChanges();

                    return Redirect("/account/rh/listardental");
                }
            }
            catch (Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep, _db, "DeleteDental, Post", user);
            }
            return Page();
        }

        public class Delete
        {
            public int Id { get; set; }
            public string UserName { get; set; }
            public string IdUsuario { get; set; }
        }
    }
}
