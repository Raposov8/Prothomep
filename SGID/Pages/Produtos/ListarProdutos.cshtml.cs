using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.Produtos
{
    [Authorize]
    public class ListarProdutosModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public List<Models.Denuo.Sb1010> DenuoProdutos { get; set; } = new List<Models.Denuo.Sb1010>();
        public List<Models.Inter.Sb1010> InterProdutos { get; set; } = new List<Models.Inter.Sb1010>();

        public ListarProdutosModel(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
        }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string Empresa,string Produto,string Anvisa,string Fabricante)
        { 
            if (Empresa == "01")
            {
                var query = ProtheusInter.Sb1010s.Where(c => c.B1Msblql != "1" && c.DELET!="*");

                if (!string.IsNullOrEmpty(Produto) && !string.IsNullOrWhiteSpace(Produto))
                {
                   query = query.Where(c => c.B1Desc.Contains(Produto));
                }

                if (!string.IsNullOrEmpty(Anvisa) && !string.IsNullOrWhiteSpace(Anvisa))
                {
                   query = query.Where(c => c.B1Reganvi.Contains(Anvisa));
                }

                if(!string.IsNullOrEmpty(Fabricante) && !string.IsNullOrWhiteSpace(Fabricante))
                {
                    query = query.Where(c => c.B1Fabric.Contains(Anvisa));
                }

                InterProdutos = query.ToList();
            }
            else
            {
                var query = ProtheusDenuo.Sb1010s.Where(c => c.B1Msblql != "1" && c.DELET != "*");

                if (!string.IsNullOrEmpty(Produto) && !string.IsNullOrWhiteSpace(Produto) )
                {
                    query = query.Where(c => c.B1Desc.Contains(Produto));
                }

                if(!string.IsNullOrEmpty(Anvisa) && !string.IsNullOrWhiteSpace(Anvisa))
                {
                    query = query.Where(c => c.B1Reganvi.Contains(Anvisa));
                }

                if (!string.IsNullOrEmpty(Fabricante) && !string.IsNullOrWhiteSpace(Fabricante))
                {
                    query = query.Where(c => c.B1Fabric.Contains(Anvisa));
                }

                DenuoProdutos = query.ToList();
            }

            return Page();
        }

    }
}
