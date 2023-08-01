using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using System.ComponentModel.DataAnnotations;

namespace SGID.Pages.Account.RH
{
    [Authorize]
    public class DetailsADMModel : PageModel
    {
        private UserManager<UserInter> _userManager;
        private ApplicationDbContext SGID;

        public Details _details { get; set; }
        public DetailsADMModel(UserManager<UserInter> userManager, ApplicationDbContext sgid)
        {
            _userManager = userManager;
            SGID = sgid;
            _details = new();
        }

        public async Task<IActionResult> OnGet(int id)
        {
            var lider = User.Identity.Name.Split("@")[0];
            var Integrante = SGID.TimeADMs.FirstOrDefault(x => x.Id == id);

            var user = await _userManager.FindByIdAsync(Integrante.IdUsuario);

            _details.Id = user.Id;
            _details.UserName = user.UserName;

            _details.Porcentagem = Integrante.Porcentagem;

            return Page();
        }

        public class Details
        {
            public string Id { get; set; }
            [Display(Name = "Integrante")]
            public string UserName { get; set; }
            [Display(Name = "% Direta")]
            public double Porcentagem { get; set; }
        }
    }
}
