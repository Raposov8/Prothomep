using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using System.ComponentModel.DataAnnotations;

namespace SGID.Pages.Account.RH
{
    public class EditarDentalModel : PageModel
    {
        private readonly ApplicationDbContext _db;
        private readonly UserManager<UserInter> _userManager;

        [BindProperty]
        public EditUserModel Editar { get; set; }

        public EditarDentalModel(ApplicationDbContext context, UserManager<UserInter> userManager)
        {
            _db = context;
            _userManager = userManager;
            Editar = new();

        }
        public async Task<IActionResult> OnGet(int id)
        {
            if (id == 0)
            {
                return LocalRedirect("Error");
            }
            var lider = User.Identity.Name.Split("@")[0];
            var usuario = _db.TimeDentals.FirstOrDefault(x => x.Id == id);

            if (usuario == null)
            {
                return LocalRedirect("Error");
            }

            Editar = new()
            {
                Id = usuario.Id,
                Email = usuario.Integrante,
                Meta = usuario.Meta,
                Porcentagem = usuario.Porcentagem,
                PorcentagemSegun = usuario.PorcentagemSeg,
                Salario = usuario.Salario,
                Teto = usuario.Teto
            };

            return Page();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            try
            {
                var lider = User.Identity.Name.Split("@")[0];
                var usuario = _db.TimeDentals.FirstOrDefault(x => x.Id == Editar.Id);

                usuario.Integrante = Editar.Email;

                usuario.Meta = Editar.Meta;
                usuario.Porcentagem = Editar.Porcentagem;
                usuario.PorcentagemSeg = Editar.PorcentagemSegun;

                usuario.Teto = Editar.Teto;
                usuario.Salario = Editar.Salario;


                _db.TimeDentals.Update(usuario);
                _db.SaveChanges();


                return Redirect("/account/rh/listardental");
            }
            catch (Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep, _db, "EditDental", user);
            }
            ModelState.AddModelError("", "Something failed.");
            return Page();
        }

        public class EditUserModel
        {
            public int Id { get; set; }

            [Required(AllowEmptyStrings = false)]
            [Display(Name = "E-mail")]
            public string Email { get; set; }

            public double Meta { get; set; }

            [Display(Name = "% Direta")]
            public double Porcentagem { get; set; }

            [Display(Name = "% Equipe")]
            public double PorcentagemSegun { get; set; } = 0.0;
            [Display(Name = "Teto")]
            public double Teto { get; set; } = 0.0;
            [Display(Name = "Salario")]
            public double Salario { get; set; } = 0.0;
        }
    }
}
