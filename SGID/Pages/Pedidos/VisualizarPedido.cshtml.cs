using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.Pedidos
{
    public class VisualizarPedidoModel : PageModel
    {
        private ApplicationDbContext DB { get; set; }
        private TOTVSDENUOContext Denuo { get; set; }
        private TOTVSINTERContext Inter { get; set; }


        public Models.Inter.Sc5010 PedidoInter { get; set; }
        public List<Models.Inter.Sc6010> ItensInter { get; set; } = new List<Models.Inter.Sc6010>();

        public Models.Denuo.Sc5010 PedidoDenuo { get; set; }

        public List<Models.Denuo.Sc6010> ItensDenuo { get; set; } = new List<Models.Denuo.Sc6010>();

        public VisualizarPedidoModel(ApplicationDbContext dB, TOTVSDENUOContext denuo, TOTVSINTERContext inter)
        {
            DB = dB;
            Denuo = denuo;
            Inter = inter;
        }

        public void OnGet(string empresa,string Id)
        {
            if (empresa == "01")
            {
                PedidoInter = Inter.Sc5010s.First(x => x.C5Num == Id);

                ItensInter = Inter.Sc6010s.Where(x => x.C6Num == Id).OrderBy(x => x.C6Item).ToList();
            }
            else
            {
                PedidoDenuo = Denuo.Sc5010s.First(x => x.C5Num == Id);

                ItensDenuo = Denuo.Sc6010s.Where(x => x.C6Num == Id).OrderBy(x=> x.C6Item).ToList();
            }
        }
    }
}
