using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;

namespace SGID.Pages.Cirurgias
{
    [Authorize(Roles = "Admin,GestorVenda,Venda,Diretoria")]
    public class NovoRecursoModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }

        private readonly IWebHostEnvironment _webHostEnvironment;

        public NovoRecursoModel(ApplicationDbContext sgid, IWebHostEnvironment webHostEnvironment, TOTVSINTERContext protheus, TOTVSDENUOContext denuo)
        {
            SGID = sgid;
            _webHostEnvironment = webHostEnvironment;
            ProtheusInter = protheus;
            ProtheusDenuo = denuo;
        }
         

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string Empresa,string Nome,string Codigo,string Cpf,string Rg,string Status,string Contratacao,string Obs,IFormFile FotoRecurso)
        {
            try
            {
                if (Empresa == "01")
                {
                    Models.Inter.Pah010 Pessoa = new()
                    {
                        PahCodins = Codigo,
                        PahNome = Nome,
                        PahRg = Rg,
                        PahCpf = Cpf,
                        PahMsblql = Status,
                        PahTipocontrato = Contratacao,
                        PahObs = Obs,
                        DELET = "",
                        PahImagem = "",
                    };

                    ProtheusInter.Pah010s.Add(Pessoa);
                    ProtheusInter.SaveChanges();

                    string Pasta = $"{_webHostEnvironment.WebRootPath}/fotos";

                    if (!Directory.Exists(Pasta))
                    {
                        Directory.CreateDirectory(Pasta);
                    }

                    if (FotoRecurso != null)
                    {
                        Pessoa.PahImagem = $"{Pessoa.RECNO}.{FotoRecurso.FileName.Split(".").Last()}";

                        string Caminho = $"{Pasta}/{Pessoa.PahImagem}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            FotoRecurso.CopyTo(fileStream);
                        }

                        ProtheusInter.Pah010s.Update(Pessoa);
                        ProtheusInter.SaveChanges();
                    }
                }
                else
                {
                    /*Models.Protheus.Pah010 Pessoa = new Models.Protheus.Pah010
                    {
                        PahCodins = Codigo,
                        PahNome = Nome,
                        PahRg = Rg,
                        PahCpf = Cpf,
                        PahMsblql = Status,
                        PahTipocontrato = Contratacao,
                        PahObs = Obs,
                        DELET = "",
                        PahImagem = "",
                    };

                    ProtheusDenuo.Pah010s.Add(Pessoa);
                    ProtheusDenuo.SaveChanges();

                    string Pasta = $"{_webHostEnvironment.WebRootPath}/fotosDenuo";

                    if (!Directory.Exists(Pasta))
                    {
                        Directory.CreateDirectory(Pasta);
                    }

                    if (FotoRecurso != null)
                    {
                        Pessoa.PahImagem = $"{Pessoa.RECNO}.{FotoRecurso.FileName.Split(".").Last()}";

                        string Caminho = $"{Pasta}/{Pessoa.PahImagem}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            FotoRecurso.CopyTo(fileStream);
                        }

                        ProtheusDenuo.Pah010s.Update(Pessoa);
                        ProtheusDenuo.SaveChanges();
                    }*/
                }   
                return LocalRedirect("/cirurgias/listarrecursos");

            }
            catch(Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoRecurso",user);

                return LocalRedirect("/error");
            }
        }
    }
}
