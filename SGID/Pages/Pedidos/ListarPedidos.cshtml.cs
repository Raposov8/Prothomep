using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Models.Denuo;
using SGID.Models.Estoque.RelatorioFaturamentoNFFab;

namespace SGID.Pages.Pedidos
{
    [Authorize]
    public class ListarPedidosModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get;set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public List<Models.Denuo.Sc5010> PedidosDenuo { get; set; } = new List<Models.Denuo.Sc5010>();
        public List<Models.Inter.Sc5010> PedidosInter { get; set; } = new List<Models.Inter.Sc5010>();
        
        public ListarPedidosModel(TOTVSDENUOContext denuo, TOTVSINTERContext inter)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
        }

        public void OnGet()
        {
        }

        public IActionResult OnPost(string NumPedido,string DTEmissao,string Empresa)
        {

            if (Empresa == "01")
            {
                var query = ProtheusInter.Sc5010s.Where(x => x.DELET != "*" && x.C5Msblql != "1");

                if (!string.IsNullOrEmpty(NumPedido) && !string.IsNullOrWhiteSpace(NumPedido))
                {
                    query = query.Where(c => c.C5Num == NumPedido);
                }

                if (!string.IsNullOrEmpty(DTEmissao) && !string.IsNullOrWhiteSpace(DTEmissao))
                {
                    query = query.Where(c => c.C5Emissao == DTEmissao);
                }

                PedidosInter = query.ToList();
            }
            else
            {
                var query = ProtheusDenuo.Sc5010s.Where(x => x.DELET != "*" && x.C5Msblql != "1");

                if (!string.IsNullOrEmpty(NumPedido) && !string.IsNullOrWhiteSpace(NumPedido))
                {
                    query = query.Where(c => c.C5Num == NumPedido);
                }

                if (!string.IsNullOrEmpty(DTEmissao) && !string.IsNullOrWhiteSpace(DTEmissao))
                {
                    query = query.Where(c => c.C5Emissao == DTEmissao);
                }

                PedidosDenuo = query.ToList();
            }

            return Page();
        }
    }
}
