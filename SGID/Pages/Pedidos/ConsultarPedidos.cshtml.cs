using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;
using SGID.Models.Inter;
using SGID.Models.Pedido;

namespace SGID.Pages.Pedidos
{
    [Authorize]
    public class ConsultarPedidosModel : PageModel
    {

        private ApplicationDbContext DB { get; set; }
        private TOTVSDENUOContext Denuo { get; set; }
        private TOTVSINTERContext Inter { get; set; }

        public List<Pedido> Pedidos { get; set; } = new List<Pedido>();

        public string Empresa { get; set; }

        public ConsultarPedidosModel(ApplicationDbContext db,TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            DB = db;
            Denuo = denuo;
            Inter = inter;
        }

        public void OnGet(string empresa)
        {
            try
            {
                Empresa = empresa;
                var date = DateTime.Now.AddYears(-4);
                var data = Convert.ToInt32(date.ToString("yyyy/MM/dd").Replace("/", ""));
                if (empresa == "01")
                {

                    Pedidos = Inter.Sc5010s.Where(x => x.C5Utpoper == "F" && x.DELET != "*" && (int)(object)x.C5XDtcir >= data)
                    .Select(x => new Pedido
                    {
                        NumeroPedido = x.C5Num,
                        Vendedor = x.C5Nomvend,
                        Medico = x.C5XNmmed,
                        Paciente = x.C5XNmpac,
                        Cliente = x.C5Nomcli,
                        DataCirurgia = x.C5XDtcir
                    }).OrderByDescending(x => x.NumeroPedido).ToList();
                }
                else
                {
                    Pedidos = Denuo.Sc5010s.Where(x => x.C5Utpoper == "F" && x.DELET != "*" && (int)(object)x.C5XDtcir >= data)
                    .Select(x => new Pedido
                    {
                        NumeroPedido = x.C5Num,
                        Vendedor = x.C5Nomvend,
                        Medico = x.C5XNmmed,
                        Paciente = x.C5XNmpac,
                        Cliente = x.C5Nomcli,
                        DataCirurgia = x.C5XDtcir
                    }).OrderByDescending(x => x.NumeroPedido).ToList();
                }
            }catch(Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, DB, "ConsultarPedidos", user);
            }
        }

    }
}
