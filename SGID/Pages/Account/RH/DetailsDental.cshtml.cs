using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;
using SGID.Data;
using System.ComponentModel.DataAnnotations;
using System.Security.Claims;

namespace SGID.Pages.Account.RH
{
    public class DetailsDentalModel : PageModel
    {
        private UserManager<UserInter> _userManager;
        private ApplicationDbContext SGID;

        public Details _details { get; set; }
        public DetailsDentalModel(UserManager<UserInter> userManager, ApplicationDbContext sgid)
        {
            _userManager = userManager;
            SGID = sgid;
            _details = new();
        }

        public async Task<IActionResult> OnGet(int id)
        {
            var lider = User.Identity.Name.Split("@")[0];
            var Integrante = SGID.TimeDentals.FirstOrDefault(x => x.Id == id);

            //var user = await _userManager.FindByIdAsync(Integrante.IdUsuario);

            _details.Id = Integrante.Id;
            _details.UserName = Integrante.Integrante;
            _details.GestorName = Integrante.Lider;

            //_details.Roles = await _userManager.GetRolesAsync(user);
            //_details.Claims = await _userManager.GetClaimsAsync(user);

            _details.Porcentagem = Integrante.Porcentagem;
            _details.Meta = Integrante.Meta;

            return Page();
        }

        public class Details
        {
            public int Id { get; set; }
            [Display(Name = "Integrante")]
            public string UserName { get; set; }
            [Display(Name = "Gestor")]
            public string GestorName { get; set; }
            public double Meta { get; set; }
            [Display(Name = "% Direta")]
            public double Porcentagem { get; set; }
            [Display(Name = "% Equipe")]
            public double PorcentagemSegun { get; set; }
            public IList<string> Roles { get; set; }
            public IList<Claim> Claims { get; set; }
        }
    }
}
