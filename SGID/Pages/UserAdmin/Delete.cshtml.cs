using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;

namespace SGID.Pages.UserAdmin
{
    [Authorize(Roles = "Admin")]
    public class DeleteModel : PageModel
    {
        private readonly UserManager<UserInter> _userManager;
        private readonly ApplicationDbContext _db;

        public Delete Input { get; set; }

        public DeleteModel(UserManager<UserInter> userManager, ApplicationDbContext db)
        {
            _userManager = userManager;
            _db = db;
            Input = new();
        }

        public async Task<ActionResult> OnGet(string id)
        {
            try
            {
                var user = await _userManager.FindByIdAsync(id);

                Input = new Delete 
                { 
                    Id = user.Id,
                    UserName = user.UserName
                };


                return Page();
            }
            catch(Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep, _db, "Delete, Get",user);
                return Redirect("/UserAdmin/ListarUsuarios/1");
            }
            
        }

        public async Task<IActionResult> OnPostAsync(string id)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    var user = await _userManager.FindByIdAsync(id);


                    var result = await _userManager.DeleteAsync(user);

                    if (!result.Succeeded)
                    {
                        ModelState.AddModelError("", result.Errors.First().Description);
                        return Page();
                    }

                    return Redirect("/UserAdmin/ListarUsuarios/1");
                }
            }
            catch(Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep,_db,"Delete Post",user);
            }
            return Page();
        }

        public class Delete 
        { 
            public string Id { get; set; }
            public string UserName { get; set; }
        }
    }
}
