using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;

namespace SGID.Pages.Account
{
    public class LoginModel : PageModel
    {
        private readonly ILogger<LoginModel> _logger;
        private readonly SignInManager<UserInter> _signInManager;
        [BindProperty]
        public InputModel Input { get; set; }

        public LoginModel(ILogger<LoginModel> logger, SignInManager<UserInter> signInManager)
        {
            _logger = logger;
            _signInManager = signInManager;
            Input = new InputModel();
        }

        public void OnGet()
        {

        }

        public async Task<IActionResult> OnPostAsync()
        {

            if (ModelState.IsValid)
            {
                // This doesn't count login failures towards account lockout
                // To enable password failures to trigger account lockout, 
                // set lockoutOnFailure: true
                var result = await _signInManager.PasswordSignInAsync(Input.Login,
                                   Input.Password, Input.RememberMe, lockoutOnFailure: true);
                if (result.Succeeded)
                {
                    _logger.LogInformation("User logged in.");
                    return LocalRedirect("/dashboard/3");
                }
                if (result.IsLockedOut)
                {
                    _logger.LogWarning("User account locked out.");
                    ModelState.AddModelError(string.Empty, "User account locked out.");
                    return Page();
                }
                else
                {
                    ModelState.AddModelError(string.Empty, "Login ou Senha incorreto.");
                    return Page();
                }
            }

            // If we got this far, something failed, redisplay form
            return Page();
        }


        public class InputModel
        {
            public string Login { get; set; } = "";
            public string Password { get; set; } = "";
            public bool RememberMe { get; set; } = false;
        }
    }
}
