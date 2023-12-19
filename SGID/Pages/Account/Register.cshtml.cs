using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using Microsoft.AspNetCore.Authorization;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Email;
using System.Net.Mail;

namespace SGID.Pages.Account
{
    [Authorize(Roles = "Admin,GestorVenda")]
    public class RegisterModel : PageModel
    {
        private readonly ILogger<RegisterModel> _logger;
        private readonly UserManager<UserInter> _userManager;
        private readonly ApplicationDbContext _db;
        private readonly IWebHostEnvironment _WEB;

        [BindProperty]
        public InputModel Input { get; set; }

        public RegisterModel(ILogger<RegisterModel> logger, UserManager<UserInter> userManager
            ,ApplicationDbContext context, IWebHostEnvironment wEB)
        {
            _logger = logger;
            _userManager = userManager;
            Input = new InputModel();
            _db = context;
            _WEB = wEB;
        }
        public void OnGet()
        {
            if (User.IsInRole("Admin"))
            {
                ViewData["RoleId"] = new SelectList(_db.Roles.Where(x => x.NormalizedName != "INSTRUMENTADOR").OrderBy(x => x.Name).ToList(), "Name", "Name");
            }
            else
            {
                ViewData["RoleId"] = new SelectList(_db.Roles.Where(x => x.NormalizedName == "VENDA").OrderBy(x => x.Name).ToList(), "Name", "Name");
            }
        }

        public async Task<IActionResult> OnPostAsync(params string[] SelectedRoles)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    var user = new UserInter { UserName = Input.Email, Email = Input.Email, EmailConfirmed = true };
                    var result = await _userManager.CreateAsync(user, Input.Password);
                    if (result.Succeeded)
                    {
                        _logger.LogInformation("User created a new account with password.");

                        if (SelectedRoles != null)
                        {
                            var result2 = await _userManager.AddToRolesAsync(user, SelectedRoles);
                            if (!result2.Succeeded)
                            {
                                foreach(var data in result2.Errors)
                                {
                                    ModelState.AddModelError("", data.Description);
                                }

                                if (User.IsInRole("Admin"))
                                {
                                    ViewData["RoleId"] = new SelectList(_db.Roles.OrderBy(x => x.Name).ToList(), "Name", "Name");
                                }
                                else
                                {
                                    ViewData["RoleId"] = new SelectList(_db.Roles.Where(x => x.NormalizedName == "VENDA").OrderBy(x => x.Name).ToList(), "Name", "Name");
                                }
                                return Page();
                            }
                        }

                        var mensagem = "";
                        var template = new
                        {
                            Titulo = "Bem-vindo(a) ao",
                            Titulo2 = "GID",
                            Email = Input.Email,
                            Senha = Input.Password,
                            Nome = Input.Email.Split("@")[0].Replace("."," ").ToUpper(),
                            Obs = ""
                        };

                        mensagem = EmailTemplate.LerArquivoHtml($"{_WEB.WebRootPath}/template/TemplateTime.html", template);

                        SmtpClient client = new SmtpClient();
                        client.Host = "smtp.office365.com";
                        client.EnableSsl = true;
                        client.Credentials = new System.Net.NetworkCredential("ti@intermedic.com.br", "interadm2018!*");
                        MailMessage mail = new MailMessage();
                        mail.Sender = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
                        mail.From = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
                        mail.To.Add(new MailAddress(Input.Email, "RECEBEDOR"));
                        mail.Bcc.Add(new MailAddress("ti@intermedic.com.br"));
                        mail.Subject = "Primeiro Acesso";
                        mail.Body = mensagem;
                        mail.IsBodyHtml = true;
                        mail.Priority = MailPriority.High;

                        client.Send(mail);

                        return LocalRedirect("/dashboard/3");
                    }
                    foreach (var error in result.Errors)
                    {
                        ModelState.AddModelError(string.Empty, error.Description);
                    }
                }
                if (User.IsInRole("Admin"))
                {
                    ViewData["RoleId"] = new SelectList(_db.Roles.OrderBy(x => x.Name).ToList(), "Name", "Name");
                }
                else
                {
                    ViewData["RoleId"] = new SelectList(_db.Roles.Where(x => x.NormalizedName == "VENDA").OrderBy(x => x.Name).ToList(), "Name", "Name");
                }
            }
            catch (Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep,_db, "Register",user);
            }
            return Page();
        }

        public class InputModel
        {
            public string Email { get; set; } = "";
            public string Password { get; set; } = "";
            public string ConfirmPassword { get; set; } = "";
        }
    }
}
