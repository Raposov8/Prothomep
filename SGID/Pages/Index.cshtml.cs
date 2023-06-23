using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;

namespace SGID.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private readonly SignInManager<UserInter> _signInManager;
        private readonly ApplicationDbContext _db;

        [BindProperty]
        public InputModel Input { get; set; }

        public IndexModel(ILogger<IndexModel> logger, 
            SignInManager<UserInter> signInManager,ApplicationDbContext context)
        {
            _logger = logger;
            _signInManager = signInManager;
            _db = context;
            Input = new InputModel();
        }

        public void OnGet()
        {

        }


        public async Task<IActionResult> OnPostAsync()
        {

            if (ModelState.IsValid)
            {

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