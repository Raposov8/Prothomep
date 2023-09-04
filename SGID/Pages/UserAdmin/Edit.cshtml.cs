using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using System.ComponentModel.DataAnnotations;

namespace SGID.Pages.UserAdmin
{
    [Authorize(Roles = "Admin")]
    public class EditModel : PageModel
    {
        private readonly ApplicationDbContext _db;
        private readonly UserManager<UserInter> _userManager;
        private readonly RoleManager<IdentityRole> _roleManager;

        [BindProperty]
        public EditUserModel Editar { get; set; }

        public EditModel(ApplicationDbContext context,UserManager<UserInter> userManager
            ,RoleManager<IdentityRole> roleManager)
        {
            _db = context;
            _userManager = userManager;
            _roleManager = roleManager;
            Editar = new();
        }
        public async Task<IActionResult> OnGet(string id)
        {
            if (id == null)
            {
                return LocalRedirect("Error");
            }
            var user = await _userManager.FindByIdAsync(id);
            if (user == null)
            {
                return LocalRedirect("Error");
            }

            var userRoles = await _userManager.GetRolesAsync(user);

            Editar = new()
            {
                Id = user.Id,
                Email = user.Email,
                RolesList = _roleManager.Roles.ToList().Select(x => new SelectListItem()
                {
                    Selected = userRoles.Contains(x.Name),
                    Text = x.Name,
                    Value = x.Name
                })
            };

            return Page();
        }

        public async Task<IActionResult> OnPostAsync(string Email,params string[] selectedRole)
        {
            try
            {
                var user = await _userManager.FindByIdAsync(Editar.Id);
                if (user == null)
                {

                    ModelState.AddModelError("", "Usuário não encontrado");
                    return Page();
                }

                user.UserName = Editar.Email;
                user.Email = Editar.Email;

                user.NormalizedEmail = Editar.Email.ToUpper();
                user.NormalizedUserName = Editar.Email.ToUpper();

                var userRoles = await _userManager.GetRolesAsync(user);

                selectedRole ??= Array.Empty<string>();

                var result = await _userManager.AddToRolesAsync(user, selectedRole.Except(userRoles).ToArray());

                if (!result.Succeeded)
                {
                    Editar = new()
                    {
                        Id = user.Id,
                        Email = user.Email,
                        RolesList = _roleManager.Roles.ToList().Select(x => new SelectListItem()
                        {
                            Selected = userRoles.Contains(x.Name),
                            Text = x.Name,
                            Value = x.Name
                        })
                    };
                    ModelState.AddModelError("", result.Errors.First().Description);
                    return Page();
                }

                result = await _userManager.RemoveFromRolesAsync(user, userRoles.Except(selectedRole).ToArray());

                if (!result.Succeeded)
                {
                    Editar = new()
                    {
                        Id = user.Id,
                        Email = user.Email,
                        RolesList = _roleManager.Roles.ToList().Select(x => new SelectListItem()
                        {
                            Selected = userRoles.Contains(x.Name),
                            Text = x.Name,
                            Value = x.Name
                        })
                    };
                    ModelState.AddModelError("", result.Errors.First().Description);
                    return Page();
                }
                return Redirect("/useradmin/listarusuarios/1");
            }
            catch (Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep, _db, "Edit", user);
            }
            ModelState.AddModelError("", "Something failed.");
            return Page();
        }

        public class EditUserModel
        {
            public string Id { get; set; }

            [Required(AllowEmptyStrings = false)]
            [Display(Name = "E-mail")]
            [EmailAddress]
            public string Email { get; set; }

            public IEnumerable<SelectListItem> RolesList { get; set; }
        }
    }
}
