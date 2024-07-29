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

        public async Task<IActionResult> OnPostAsync()
        {
            try
            {
                if (ModelState.IsValid)
                {
                    var integrante = Input.Email;

                        var time = new TimeDental
                        {
                            DataCriacao = DateTime.Now,
                            Status = true,
                            Lider = "MARCOS.PARRA",
                            Integrante = integrante.Split("@")[0],
                            Meta = Input.Meta,
                            Porcentagem = Input.Porcentagem,
                            PorcentagemSeg = Input.PorcentagemSegun,
                            IdUsuario = "",
                            Salario = Input.Salario,
                            Garantia = Input.Garantia,
                            AtingimentoMeta = Input.AtingimentoMeta,
                            PorcentagemEtapaDois = Input.PorcentagemMeta2,
                            PorcentagemEtapaUm = Input.PorcentagemMeta1,
                            TipoComissao = Input.TipoComissao
                        };

                        _db.TimeDentals.Add(time);
                        _db.SaveChanges();

                        return LocalRedirect("/account/rh/listardental");
                    
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
            public double Meta { get; set; } = 0.0;
            [Display(Name = "% Direta")]
            public double Porcentagem { get; set; } = 0.0;
            [Display(Name = "% Equipe")]
            public double PorcentagemSegun { get; set; } = 0.0;
            [Display(Name = "Teto")]
            public double Teto { get; set; } = 0.0;
            [Display(Name = "Salario")]
            public double Salario { get; set; } = 0.0;
            [Display(Name = "Garantia")]
            public double Garantia { get; set; } = 0.0;
            [Display(Name = "Atingimento Meta")]
            public double AtingimentoMeta { get; set; } = 0.0;
            [Display(Name = "% Etapa 1")]
            public double PorcentagemMeta1 { get; set; } = 0.0;
            [Display(Name = "% Etapa 2")]
            public double PorcentagemMeta2 { get; set; } = 0.0;
            [Display(Name = "Tipo Comissão")]
            public string? TipoComissao { get; set; } = "";
        }
    }
}

