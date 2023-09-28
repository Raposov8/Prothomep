using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using System.ComponentModel.DataAnnotations;

namespace SGID.Pages.Account.RH
{
    public class CadastroDentalModel : PageModel
    {

        private readonly ILogger<RegisterModel> _logger;
        private readonly UserManager<UserInter> _userManager;
        private readonly ApplicationDbContext _db;
        private readonly IWebHostEnvironment _WEB;

        [BindProperty]
        public InputModel Input { get; set; }

        public CadastroDentalModel(ILogger<RegisterModel> logger, UserManager<UserInter> userManager
            , ApplicationDbContext context, IWebHostEnvironment wEB)
        {
            _logger = logger;
            _userManager = userManager;
            Input = new InputModel();
            _db = context;
            _WEB = wEB;

        }
        public void OnGet()
        {

        }

        public async Task<IActionResult> OnPostAsync(params string[] SelectedLinhas)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    //var user = new UserInter { UserName = Input.Email, Email = Input.Email, EmailConfirmed = true };
                    //var usuariop = await _userManager.FindByEmailAsync(user.Email);



                    //var lider = Input.GestorEmail.Split("@")[0];
                    var integrante = Input.Email;//.Split("@")[0];

                        var time = new TimeDental
                        {
                            DataCriacao = DateTime.Now,
                            Status = true,
                            Lider = "MARCOS.PARRA",
                            Integrante = integrante,
                            Meta = Input.Meta,
                            Porcentagem = Input.Porcentagem,
                            PorcentagemSeg = Input.PorcentagemSegun,
                            //IdUsuario = usuariop.Id
                        };

                        _db.TimeDentals.Add(time);
                        _db.SaveChanges();

                        return LocalRedirect("/account/rh/listartime");
                    
                }
            }
            catch (Exception excep)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(excep, _db, "CadastroTimeDental", user);
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
        }
    }
}

