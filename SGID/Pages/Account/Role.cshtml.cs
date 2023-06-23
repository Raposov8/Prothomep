using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;

namespace SGID.Pages.Account
{
    [Authorize(Roles="Admin")]
    public class RoleModel : PageModel
    {
        private readonly RoleManager<IdentityRole> _roleManager;
        private ApplicationDbContext SGID { get; set; }

        public RoleModel(RoleManager<IdentityRole> roleManager,ApplicationDbContext Sgid)
        {
            _roleManager = roleManager;
            SGID = Sgid;
            Input = new InputModel();
        }

        [BindProperty]
        public InputModel Input { get; set; }

        public void OnGet()
        {
        }

        
        public async Task<IActionResult> OnPostAsync()
        {
            try
            {
                if (ModelState.IsValid)
                {
                    var role = new IdentityRole(Input.Name);
                    var roleresult = await _roleManager.CreateAsync(role);
                    if (!roleresult.Succeeded)
                    {
                        ModelState.AddModelError("", roleresult.Errors.First().Description);
                        return Page();
                    }
                    return Redirect("/dashboard/3");
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Role",user);
            }

            return Page();
        }



        public class InputModel
        {
            [Required]
            public string Name { get; set; }
        }
    }
}
