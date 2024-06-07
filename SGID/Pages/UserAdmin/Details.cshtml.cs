using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;
using System.Security.Claims;

namespace SGID.Pages.UserAdmin
{
    [Authorize(Roles = "Admin")]
    public class DetailsModel : PageModel
    {
        private UserManager<UserInter> _userManager;

        public Details _details { get; set; }
        public DetailsModel(UserManager<UserInter> userManager)
        {
            _userManager = userManager;
            _details = new();
        }

        public async Task<IActionResult> OnGet(string id)
        {
            var user = await _userManager.FindByIdAsync(id);

            _details.Id = user.Id;
            _details.UserName = user.UserName;

            _details.Roles = await _userManager.GetRolesAsync(user);
            _details.Claims = await _userManager.GetClaimsAsync(user);

            return Page();
        }

        public class Details
        {
            public string Id { get; set; }
            public string UserName { get; set; }
            public IList<string> Roles { get; set; }
            public IList<Claim> Claims { get; set; }
        }
    }
}
