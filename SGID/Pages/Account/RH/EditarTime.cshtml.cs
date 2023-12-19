using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;
using System.ComponentModel.DataAnnotations;

namespace SGID.Pages.Account.RH
{
    [Authorize(Roles = "Admin,RH,Diretoria")]
    public class EditarTimeModel : PageModel
    {
        private readonly ApplicationDbContext _db;
        private readonly UserManager<UserInter> _userManager;
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        [BindProperty]
        public EditUserModel Editar { get; set; }
        public List<string> Linhas { get; set; } = new List<string>();
        public List<string> Gerente { get; set; } = new List<string> { "SIM", "NAO" };
        public List<string> Regioes { get; set; } = new List<string> { "CAPITAL", "INTERIOR" };

        public EditarTimeModel(ApplicationDbContext context, UserManager<UserInter> userManager
            ,TOTVSDENUOContext DENUO, TOTVSINTERContext INTER)
        {
            _db = context;
            _userManager = userManager;
            Editar = new();
            ProtheusDenuo = DENUO;
            ProtheusInter = INTER;
        }
        public async Task<IActionResult> OnGet(int id)
        {
            if (id == 0)
            {
                return LocalRedirect("Error");
            }
            var lider = User.Identity.Name.Split("@")[0];
            var usuario = _db.Times.FirstOrDefault(x => x.Id == id);

            if (usuario == null)
            {
                return LocalRedirect("Error");
            }

            var user = await _userManager.FindByIdAsync(usuario.IdUsuario);
            if (user == null)
            {
                return LocalRedirect("Error");
            }

            var linhasInter = ProtheusInter.Sa3010s.Where(x => x.A3Xdescun != "").Select(x => x.A3Xdescun.Trim()).Distinct().ToList();
            var linhasDenuo = ProtheusDenuo.Sa3010s.Where(x => x.A3Xdescun != "").Select(x => x.A3Xdescun.Trim()).Distinct().ToList();

            Linhas = linhasInter.Union(linhasDenuo).ToList();

            var Produtos = _db.TimeProdutos.Where(x => x.TimeId == usuario.Id).Select(x=>x.Produto).ToList();

            var email = user.Email.Split("@")[1];
            Editar = new()
            {
                Id = user.Id,
                Email = user.Email,
                Meta = usuario.Meta,
                EmailGestor = $"{usuario.Lider}@{email}",
                Porcentagem = usuario.Porcentagem,
                PorcentagemSegun = usuario.PorcentagemSeg,
                GerenProd = usuario.GerenProd,
                PorcentagemProd = usuario.PorcentagemGenProd,
                RolesList = Linhas.Select(x => new SelectListItem()
                {
                    Selected = Produtos.Contains(x),
                    Text = x,
                    Value = x
                }),
                Teto = usuario.Teto,
                Salario = usuario.Salario,
                Garantia = usuario.Garantia
                //Setor = usuario.TipoFaturamento
            };

            return Page();
        }

        public async Task<IActionResult> OnPostAsync(string GerenProd = "0", params string[] SelectedLinhas)
        {
            try
            {
                var lider = User.Identity.Name.Split("@")[0];
                var usuario = _db.Times.FirstOrDefault(x => x.IdUsuario == Editar.Id);

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

                usuario.Lider = Editar.EmailGestor.Split("@")[0];
                usuario.Integrante = Editar.Email.Split("@")[0];

                usuario.Meta = Editar.Meta;
                usuario.Porcentagem = Editar.Porcentagem;
                usuario.PorcentagemSeg = Editar.PorcentagemSegun;
                usuario.GerenProd = GerenProd;
                usuario.PorcentagemGenProd = Editar.PorcentagemProd;
                usuario.Teto = Editar.Teto;
                usuario.Salario = Editar.Salario;
                usuario.Garantia = Editar.Garantia;
                //usuario.TipoFaturamento = Editar.Setor;

                _db.Times.Update(usuario);
                _db.SaveChanges();

                var removeProduto = new List<string>(Linhas.Except(SelectedLinhas).ToArray());

                removeProduto.ForEach(prod =>
                {
                    var Produto = _db.TimeProdutos.FirstOrDefault(x => x.TimeId == usuario.Id && x.Produto == prod.Trim());

                    if (Produto != null)
                    {
                        _db.Remove(Produto);
                        _db.SaveChanges();
                    }
                });

                var addProduto = new List<string>(SelectedLinhas);

                addProduto.ForEach(prod =>
                {
                    var Produto = _db.TimeProdutos.FirstOrDefault(x => x.TimeId == usuario.Id && x.Produto == prod.Trim());

                    if (Produto == null)
                    {
                        Produto = new TimeProduto
                        {
                            TimeId = usuario.Id,
                            Produto = prod.Trim()
                        };

                        _db.Add(Produto);
                        _db.SaveChanges();
                    }
                });


                return Redirect("/account/rh/listartime");
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
            [Display(Name = "E-mail Gestor")]
            [EmailAddress]
            public string EmailGestor { get; set; }
            
            [Required(AllowEmptyStrings = false)]
            [Display(Name = "E-mail")]
            [EmailAddress]
            public string Email { get; set; }

            public double Meta { get; set; }

            [Display(Name = "% Direta")]
            public double Porcentagem { get; set; }

            [Display(Name = "% Equipe")]
            public double PorcentagemSegun { get; set; } = 0.0;
            [Display(Name = "% Sobre Linha")]
            public double PorcentagemProd { get; set; } = 0.0;
            [Display(Name = "É Gestor de Produtos?")]
            public string GerenProd { get; set; }

            [Display(Name = "Teto")]
            public double Teto { get; set; } = 0.0;
            [Display(Name = "Interior ou Capital?")]
            public string TipoVendedor { get; set; }
            [Display(Name = "Salario")]
            public double Salario { get; set; } = 0.0;
            [Display(Name = "Garantia")]
            public double Garantia { get; set; } = 0.0;

            //[Display(Name = "Setor Comercial")]
            //public string Setor { get; set; }
            public IEnumerable<SelectListItem> RolesList { get; set; }
        }
    }
}
