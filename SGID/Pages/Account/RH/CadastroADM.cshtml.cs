using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;
using SGID.Models.Email;
using System.ComponentModel.DataAnnotations;
using System.Net.Mail;

namespace SGID.Pages.Account.RH
{
    [Authorize]
    public class CadastroADMModel : PageModel
    {
        private readonly ILogger<RegisterModel> _logger;
        private readonly UserManager<UserInter> _userManager;
        private readonly ApplicationDbContext _db;
        private readonly IWebHostEnvironment _WEB;

        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        [BindProperty]
        public InputModel Input { get; set; }


        public CadastroADMModel(ILogger<RegisterModel> logger, UserManager<UserInter> userManager
            , ApplicationDbContext context, IWebHostEnvironment wEB, TOTVSINTERContext INTER, TOTVSDENUOContext DENUO)
        {
            _logger = logger;
            _userManager = userManager;
            Input = new InputModel();
            _db = context;
            _WEB = wEB;
            ProtheusInter = INTER;
            ProtheusDenuo = DENUO;
        }
        public void OnGet()
        {

        }

        public async Task<IActionResult> OnPostAsync()
        {
            try
            {
                if (ModelState.IsValid)
                {
                    var user = new UserInter { UserName = Input.Email, Email = Input.Email, EmailConfirmed = true };
                    user.UsuarioCriacao = User.Identity.Name.Split("@")[0];
                    user.CriacaoDate = DateTime.Now;
                    var usuariop = await _userManager.FindByEmailAsync(user.Email);
                    if (usuariop == null)
                    {
                        var result = await _userManager.CreateAsync(user, "GID@Acess@");
                        if (result.Succeeded)
                        {
                            _logger.LogInformation("User created a new account with password.");

                            var integrante = Input.Email.Split("@")[0];

                            var usuario = await _userManager.FindByEmailAsync(Input.Email);

                            var time = new TimeADM
                            {
                                DataCriacao = DateTime.Now,
                                Status = true,
                                Integrante = integrante,
                                Porcentagem = Input.Porcentagem,
                                IdUsuario = usuario.Id,
                                Salario = Input.Salario,
                                Teto = Input.Teto
                            };

                            _db.TimeADMs.Add(time);
                            _db.SaveChanges();

                            var mensagem = "";
                            var template = new
                            {
                                Titulo = "Bem-vindo(a) ao",
                                Titulo2 = "GID",
                                Email = Input.Email,
                                Senha = "GID@Acess@",
                                Nome = integrante,
                                Obs = ""
                            };

                            mensagem = EmailTemplate.LerArquivoHtml($"{_WEB.WebRootPath}/template/TemplateTime.html", template);

                            SmtpClient client = new SmtpClient();
                            client.Host = "smtp.office365.com";
                            client.EnableSsl = true;
                            client.Credentials = new System.Net.NetworkCredential("ti@flowmeds.com.br", "interadm2018!*");
                            MailMessage mail = new MailMessage();
                            mail.Sender = new MailAddress("ti@flowmeds.com.br", "ENVIADOR");
                            mail.From = new MailAddress("ti@flowmeds.com.br", "ENVIADOR");
                            mail.To.Add(new MailAddress(Input.Email, "RECEBEDOR"));
                            mail.CC.Add(new MailAddress("ti@flowmeds.com.br"));
                            mail.Subject = "Solicita��o de Acesso";
                            mail.Body = mensagem;
                            mail.IsBodyHtml = true;
                            mail.Priority = MailPriority.High;

                            client.Send(mail);

                            return LocalRedirect("/account/rh/listaradm");
                        }
                        foreach (var error in result.Errors)
                        {
                            ModelState.AddModelError(string.Empty, error.Description);
                        }
                    }
                    else
                    {

                        var integrante = Input.Email.Split("@")[0];

                        var time = new TimeADM
                        {
                            DataCriacao = DateTime.Now,
                            Status = true,
                            Integrante = integrante,
                            Porcentagem = Input.Porcentagem,
                            IdUsuario = usuariop.Id,
                            Salario = Input.Salario,
                            Teto = Input.Teto
                        };

                        _db.TimeADMs.Add(time);
                        _db.SaveChanges();


                        return LocalRedirect("/account/rh/listaradm");
                    }
                }
            }
            catch (Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep, _db, "CadastroADM", user);
            }
            return Page();
        }

        public class InputModel
        {
            [Display(Name = "E-mail")]
            public string Email { get; set; } = "";
            [Display(Name = "% Sobre Faturamento")]
            public double Porcentagem { get; set; } = 0.0;
            [Display(Name = "Teto")]
            public double Teto { get; set; } = 0.0;
            [Display(Name = "Salario")]
            public double Salario { get; set; } = 0.0;
        }
    }
}
