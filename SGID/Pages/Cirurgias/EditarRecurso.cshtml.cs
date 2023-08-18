using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;

namespace SGID.Pages.Cirurgias
{
    [Authorize(Roles = "Admin,GestorVenda,Venda,Diretoria")]
    public class EditarRecursoModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        private TOTVSINTERContext Protheus { get; set; }

        private readonly IWebHostEnvironment _webHostEnvironment;
        public Pah010 Instrumentador { get; set; }

        public EditarRecursoModel(ApplicationDbContext sgid, IWebHostEnvironment webHostEnvironment, TOTVSINTERContext protheus)
        {
            SGID = sgid;
            _webHostEnvironment = webHostEnvironment;
            Protheus = protheus;
        }


        public void OnGet(int Id)
        {
            try
            {
                Instrumentador = Protheus.Pah010s.FirstOrDefault(x => x.RECNO == Id);
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "EditarRecurso Get",user);
            }
        }

        public IActionResult OnPostAsync(int Id,string Nome, string Codigo, string Cpf, string Rg, string Status, string Contratacao, string Obs, IFormFile FotoRecurso)
        {
            try
            {
                var Recurso = Protheus.Pah010s.First(x => x.RECNO == Id);

                Recurso.PahNome = Nome;
                Recurso.PahCodins = Codigo;
                Recurso.PahCpf = Cpf;
                Recurso.PahRg = Rg;
                Recurso.PahMsblql = Status;
                Recurso.PahTipocontrato = Contratacao;
                Recurso.PahObs = Obs;

                if (FotoRecurso != null)
                {
                    string Pasta = $"{_webHostEnvironment.WebRootPath}/fotos";

                    if (!Directory.Exists(Pasta))
                    {
                        Directory.CreateDirectory(Pasta);
                    }

                    string Caminho = $"{Pasta}/{Recurso.RECNO}.{FotoRecurso.FileName.Split(".")[1]}";
                    using Stream fileStream = new FileStream(Caminho, FileMode.Create);
                    FotoRecurso.CopyTo(fileStream);
                }

                Protheus.Pah010s.Update(Recurso);
                Protheus.SaveChanges();
                return LocalRedirect("/cirurgias/listarrecursos");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "EditarRecurso",user);

                return LocalRedirect("/error");
            }
        }
    }
}
