using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Denuo;
using SGID.Models.Financeiro;

namespace SGID.Pages.Relatorios.Financeiro
{
    public class AReceberModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }

        public List<AReceber> Relatorio { get; set; } = new List<AReceber>(); 

        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public AReceberModel(TOTVSDENUOContext protheus)
        {
            Protheus = protheus;
        }

        public void OnGet()
        {
        }

        public IActionResult OnPost(DateTime DataInicio, DateTime DataFim)
        {

            Inicio = DataInicio;
            Fim = DataFim;

            var query = (from SE1 in Protheus.Se1010s
                         join SA3 in Protheus.Sa3010s on SE1.E1Vend1 equals SA3.A3Cod
                         where Convert.ToInt32(SE1.E1Baixa) >= Convert.ToInt32(DataInicio.ToString("yyyy/MM/dd").Replace("/", "")) 
                         && Convert.ToInt32(SE1.E1Baixa) <= Convert.ToInt32(DataFim.ToString("yyyy/MM/dd").Replace("/", ""))
                         select new AReceber
                         {
                            Num = SE1.E1Num,
                            Parcela = SE1.E1Parcela,
                            Tipo = SE1.E1Tipo,
                            Naturez = SE1.E1Naturez,
                            NomCli = SE1.E1Nomcli,
                            Emissao = SE1.E1Emissao,
                            Valor = SE1.E1Valor,
                            Baixa = SE1.E1Baixa,
                            NomeVende = SA3.A3Nome,
                            Login = SA3.A3Xlogin
                         }
                         ).ToList();


            return Page();
        }
    }
}
