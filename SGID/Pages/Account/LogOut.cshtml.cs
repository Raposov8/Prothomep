using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;

namespace SGID.Pages.Account
{
    public class LogOutModel : PageModel
    {
        private readonly SignInManager<UserInter> _signInManager;
        private readonly ILogger<LogOutModel> _logger;

        public LogOutModel(SignInManager<UserInter> signInManager, ILogger<LogOutModel> logger)
        {
            _signInManager = signInManager;
            _logger = logger;
        }

        public async Task<IActionResult> OnGet()
        {
            await _signInManager.SignOutAsync();
            _logger.LogInformation("User logged out.");

            return RedirectToPage("/index");
        }
    }
}
