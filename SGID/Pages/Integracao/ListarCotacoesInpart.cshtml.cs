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

        public ListarCotacoesInpartModel(TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
        }

        public void OnGet()
        {

            Cotacoes = new IntegracaoInPart().ListarCotacoes().Result;
        }


        public JsonResult OnGetCotacao(int id)
        {

            Cotacoes = new IntegracaoInPart().ListarCotacoes().Result;

            var cotacao = Cotacoes.FirstOrDefault(x=> x.idCotacao == id);

            return new JsonResult(cotacao);
        }

        public IActionResult OnPost()
        {
            return LocalRedirect("");
        }
    }
}
