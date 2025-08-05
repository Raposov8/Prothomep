using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;
using SGID.Data;
using Microsoft.AspNetCore.Authorization;
using SGID.Data.Models;
using SGID.Models.Email;
using System.Net.Mail;
using System.ComponentModel.DataAnnotations;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.Account.RH
{
    [Authorize(Roles = "Admin,RH,Diretoria")]
    public class CadastroTimeModel : PageModel
    {
        private readonly ILogger<RegisterModel> _logger;
        private readonly UserManager<UserInter> _userManager;
        private readonly ApplicationDbContext _db;
        private readonly IWebHostEnvironment _WEB;

        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        [BindProperty]
        public InputModel Input { get; set; }

        public List<string> Linhas { get; set; } = new List<string>();

        public List<string> Gerente { get; set; } = new List<string> { "SIM", "NAO" };

        public List<string> Regioes { get; set; } = new List<string> { "CAPITAL", "INTERIOR" };

        public CadastroTimeModel(ILogger<RegisterModel> logger, UserManager<UserInter> userManager
            , ApplicationDbContext context, IWebHostEnvironment wEB,TOTVSINTERContext INTER,TOTVSDENUOContext DENUO)
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

            var linhasInter = ProtheusInter.Sa3010s.Where(x => x.A3Xdescun != "" && x.A3Xdescun != "DENTAL                        ").Select(x => x.A3Xdescun).Distinct().ToList();
            var linhasDenuo = ProtheusDenuo.Sa3010s.Where(x => x.A3Xdescun != "" && x.A3Xdescun != "DENTAL                        ").Select(x => x.A3Xdescun).Distinct().ToList();

            Linhas = linhasInter.Union(linhasDenuo).ToList();
        }

        public async Task<IActionResult> OnPostAsync( params string[] SelectedLinhas)
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

                            var lider = Input.GestorEmail.Split("@")[0];
                            var integrante = Input.Email.Split("@")[0];

                            var usuario = await _userManager.FindByEmailAsync(Input.Email);

                            var time = new Time
                            {
                                DataCriacao = DateTime.Now,
                                Status = true,
                                Lider = lider,
                                Integrante = integrante,
                                Meta = Input.Meta,
                                Porcentagem = Input.Porcentagem,
                                PorcentagemSeg = Input.PorcentagemSegun,
                                GerenProd = Input.GerenProd,
                                PorcentagemGenProd = Input.PorcentagemProd,
                                IdUsuario = usuario.Id,
                                TipoVendedor = Input.TipoVendedor,
                                Teto = Input.Teto,
                                Salario = Input.Salario,
                                TipoFaturamento = Input.Setor,
                                Garantia = Input.Garantia

                            };

                            _db.Times.Add(time);
                            _db.SaveChanges();

                            var addProduto = new List<string>(SelectedLinhas);

                            addProduto.ForEach(prod =>
                            {
                                var Produto = _db.TimeProdutos.FirstOrDefault(x => x.TimeId == time.Id && x.Produto == prod.Trim());

                                if (Produto == null)
                                {
                                    Produto = new TimeProduto
                                    {
                                        TimeId = time.Id,
                                        Produto = prod.Trim()
                                    };

                                    _db.Add(Produto);
                                    _db.SaveChanges();
                                }
                            });



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
                            client.Credentials = new System.Net.NetworkCredential("ti@prothomep.com.br", "interadm2018!*");
                            MailMessage mail = new MailMessage();
                            mail.Sender = new MailAddress("ti@prothomep.com.br", "ENVIADOR");
                            mail.From = new MailAddress("ti@prothomep.com.br", "ENVIADOR");
                            mail.To.Add(new MailAddress(Input.Email, "RECEBEDOR"));
                            mail.CC.Add(new MailAddress("ti@prothomep.com.br"));
                            mail.Subject = "Solicita��o de Acesso";
                            mail.Body = mensagem;
                            mail.IsBodyHtml = true;
                            mail.Priority = MailPriority.High;

                            client.Send(mail);

                            return LocalRedirect("/account/rh/listartime");
                        }
                        foreach (var error in result.Errors)
                        {
                            ModelState.AddModelError(string.Empty, error.Description);
                        }
                    }
                    else
                    {

                        var lider = Input.GestorEmail.Split("@")[0];
                        var integrante = Input.Email.Split("@")[0];

                        var time = new Time
                        {
                            DataCriacao = DateTime.Now,
                            Status = true,
                            Lider = lider,
                            Integrante = integrante,
                            Meta = Input.Meta,
                            Porcentagem = Input.Porcentagem,
                            PorcentagemSeg = Input.PorcentagemSegun,
                            GerenProd = Input.GerenProd,
                            PorcentagemGenProd = Input.PorcentagemProd,
                            IdUsuario = usuariop.Id,
                            TipoVendedor = Input.TipoVendedor,
                            Teto = Input.Teto,
                            Salario = Input.Salario,
                            TipoFaturamento = Input.Setor,
                            Garantia = Input.Garantia

                        };

                        _db.Times.Add(time);
                        _db.SaveChanges();

                        var addProduto = new List<string>(SelectedLinhas);

                        addProduto.ForEach(prod =>
                        {
                            var Produto = _db.TimeProdutos.FirstOrDefault(x => x.TimeId == time.Id && x.Produto == prod.Trim());

                            if (Produto == null)
                            {
                                Produto = new TimeProduto
                                {
                                    TimeId = time.Id,
                                    Produto = prod.Trim()
                                };

                                _db.Add(Produto);
                                _db.SaveChanges();
                            }
                        });

                        return LocalRedirect("/account/rh/listartime");
                    }
                }
            }
            catch (Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep, _db, "CadastroTime", user);
            }
            return Page();
        }

        public class InputModel
        {
            [Display(Name = "E-mail")]
            public string Email { get; set; } = "";
            [Display(Name = "E-mail Gestor")]
            public string GestorEmail { get; set; } = "";
            public double Meta { get; set; } = 0.0;
            [Display(Name = "% Direta")]
            public double Porcentagem { get; set; } = 0.0;
            [Display(Name = "% Equipe")]
            public double PorcentagemSegun { get; set; } = 0.0;
            [Display(Name = "% Sobre Linha")]
            public double PorcentagemProd { get; set; } = 0.0;
            [Display(Name = "É Gestor de Produtos")]
            public string GerenProd { get; set; }
            [Display(Name = "Teto")]
            public double Teto { get; set; } = 0.0;
            [Display(Name = "Interior ou Capital?")]
            public string TipoVendedor { get; set; }
            [Display(Name = "Salario")]
            public double Salario { get; set; } = 0.0;
            [Display(Name = "Setor Comercial")]
            public string Setor { get; set; }
            [Display(Name = "Garantia")]
            public double Garantia { get; set; } = 0.0;
        }
    }

}
