using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;

namespace SGID.Pages.Instrumentador
{
    public class EditarInstrumentadorModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        public RecursoIntrumentador Instrumentador { get; set; }

        public EditarInstrumentadorModel(ApplicationDbContext sgid)
        {
            SGID = sgid;
        }

        public void OnGet(int Id)
        {
            Instrumentador = SGID.Instrumentadores.FirstOrDefault(x => x.Id == Id);
        }


        public IActionResult OnPost(int Id,string Nome, DateTime DtNascimento, string Endereco, string Municipio, string Bairro,
            string Estado, string Pais, string Email, string Telefone1, string Telefone2, string CBO, string RG, string CPF,
            string PisPasep, string TipoChave, string ChavePix, string Banco, string AG, string CC, string Remuneracao, string ServCon,
            string NomeEmpresa, string CNPJ, string Tipo)
        {
            try
            {

                Instrumentador = SGID.Instrumentadores.FirstOrDefault(x => x.Id == Id);


                Instrumentador.Nome = Nome;
                Instrumentador.Nascimento = DtNascimento;
                Instrumentador.Endereco = Endereco;
                Instrumentador.Municipio = Municipio;
                Instrumentador.Bairro = Bairro;
                Instrumentador.Estado = Estado;
                Instrumentador.Pais = Pais;
                Instrumentador.Email = Email;
                Instrumentador.Telefone = Telefone1;
                Instrumentador.Telefone2 = Telefone2;
                Instrumentador.CBO = CBO;
                Instrumentador.RG = RG;
                Instrumentador.CPF = CPF;
                Instrumentador.PISPASEP = PisPasep;
                Instrumentador.TipoChave = TipoChave;
                Instrumentador.ChavePix = ChavePix;
                Instrumentador.Banco = Banco;
                Instrumentador.Ag = AG;
                Instrumentador.CC = CC;
                Instrumentador.Remuneracao = Remuneracao;
                Instrumentador.ServCon = ServCon;
                Instrumentador.NomeEmpresa = NomeEmpresa;
                Instrumentador.CNPJ = CNPJ;
                Instrumentador.Tipo = Tipo;



                SGID.Instrumentadores.Update(Instrumentador);
                SGID.SaveChanges();

                return LocalRedirect("/DashBoards/DashboardGestorInstrumentador/0");
            }
            catch (Exception E)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(E, SGID, "EditarInstrumentador", user);

                return LocalRedirect("/error");
            }
        }
    }
}
