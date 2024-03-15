using Intergracoes.Inpart;
using Intergracoes.Inpart.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.Integracao
{
    [Authorize]
    public class ListarCotacoesInpartModel : PageModel
    {
        private TOTVSINTERContext ProtheusInter { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }

        public List<Cotacao> Cotacoes { get; set; }

        public int Pagina { get; set; }
        public string Empresa { get; set; }

        public ListarCotacoesInpartModel(TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
        }

        public void OnGet(string empresa)
        {
            Pagina = 1;
            Empresa = empresa;

            Cotacoes = new IntegracaoInPart().ListarCotacoes(empresa,0).Result;
        }


        public JsonResult OnGetCotacao(int id,string empresa,int skip)
        {

            Cotacoes = new IntegracaoInPart().ListarCotacoes(empresa,skip-1).Result;

            var cotacao = Cotacoes.FirstOrDefault(x=> x.idCotacao == id);

            return new JsonResult(cotacao);
        }

        public IActionResult OnPost(string empresa,int skip)
        {
            Pagina = skip;
            Empresa = empresa;

            Cotacoes = new IntegracaoInPart().ListarCotacoes(empresa, skip-1).Result;

            return Page();
        }
    }
}
