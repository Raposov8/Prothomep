using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Controladoria;
using SGID.Models.Denuo;
using SGID.Models.Inter;
using SGID.Models.Qualidade;

namespace SGID.Pages.DashBoards
{
    public class DashBoardQualidadeModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public List<ProdutoValidade> Produtos { get; set; } = new List<ProdutoValidade>();

        public DashBoardQualidadeModel(TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
        }
        public void OnGet()
        {
            var data = DateTime.Now;

            var DataFim = Convert.ToInt32($"{data.Year}{data.Month:D2}{data.Day:D2}");

            var resultadoInter = ProtheusInter.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1" && (int)(object)x.B1Dtvalid < DataFim && x.B1Dtvalid != "")
                .Select(x => new ProdutoValidade
                {
                    Codigo = x.B1Cod,
                    Empresa = "INTERMEDIC",
                    Descricao = x.B1Desc,
                    DtValidade = Convert.ToDateTime($"{x.B1Dtvalid.Substring(4, 2)}/{x.B1Dtvalid.Substring(6, 2)}/{x.B1Dtvalid.Substring(0, 4)}"),
                }).ToList();

            var resultadoDenuo = ProtheusDenuo.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1" && (int)(object)x.B1Dtvalid < DataFim && x.B1Dtvalid != "")
                .Select(x => new ProdutoValidade
                {
                    Codigo = x.B1Cod,
                    Empresa = "DENUO",
                    Descricao = x.B1Desc,
                    DtValidade = Convert.ToDateTime($"{x.B1Dtvalid.Substring(4, 2)}/{x.B1Dtvalid.Substring(6, 2)}/{x.B1Dtvalid.Substring(0, 4)}"),
                }).ToList();


            Produtos = resultadoDenuo.Concat(resultadoInter).OrderByDescending(x=> x.DtValidade).ToList();
        }
    }
}
