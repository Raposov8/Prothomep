using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;

namespace SGID.Pages.Instrumentador
{
    public class NovoInstrumentadorModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        public NovoInstrumentadorModel(ApplicationDbContext sgid) => SGID = sgid; 

        public void OnGet()
        {
        }

        public IActionResult OnPost(string Nome,DateTime DtNascimento, string Endereco,string Municipio,string Bairro,
            string Estado,string Pais,string Email,string Telefone1,string Telefone2,string CBO,string RG,string CPF,
            string PisPasep,string TipoChave,string ChavePix,string Banco,string AG,string CC,string Remuneracao,string ServCon)
        {
            try
            {

                var Instrumentador = new RecursoIntrumentador
                { 
                    Nome = Nome,
                    Nascimento = DtNascimento,
                    Endereco = Endereco,
                    Municipio = Municipio,
                    Bairro = Bairro,
                    Estado = Estado,
                    Pais = Pais,
                    Email = Email,
                    Telefone = Telefone1,
                    Telefone2 = Telefone2,
                    CBO = CBO,
                    RG = RG,
                    CPF = CPF,
                    PISPASEP = PisPasep,
                    TipoChave = TipoChave,
                    ChavePix = ChavePix,
                    Banco = Banco,
                    Ag = AG,
                    CC = CC,
                    Remuneracao = Remuneracao,
                    ServCon = ServCon
                };

                SGID.Instrumentadores.Add(Instrumentador);
                SGID.SaveChanges();

                return LocalRedirect("/DashBoards/DashBoardGestorInstrumentador");
            }
            catch(Exception E)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(E, SGID, "NovoInstrumentador", user);

                return LocalRedirect("/error");
            }
        }
    }
}
