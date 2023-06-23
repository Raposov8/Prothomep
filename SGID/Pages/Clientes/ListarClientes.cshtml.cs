using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.Clientes
{
    [Authorize]
    public class ListarClientesModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public List<Models.Denuo.Sa1010> ClientesDenuo { get; set; } = new List<Models.Denuo.Sa1010>();
        public List<Models.Inter.Sa1010> ClientesInter { get; set; } = new List<Models.Inter.Sa1010>();

        public ListarClientesModel(TOTVSDENUOContext denuo, TOTVSINTERContext inter)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
        }
        public void OnGet()
        {
        }

        public IActionResult OnPost(string Cliente,string Empresa)
        {

            if (Empresa == "01")
            {
                var query = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Msblql != "1");

                if (!string.IsNullOrEmpty(Cliente) && !string.IsNullOrWhiteSpace(Cliente))
                {
                    query = query.Where(c => c.A1Nome.Contains(Cliente) || c.A1Nreduz.Contains(Cliente));
                }

                ClientesInter = query.ToList();
            }
            else
            {
                var query = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Msblql != "1");

                if (!string.IsNullOrEmpty(Cliente) && !string.IsNullOrWhiteSpace(Cliente))
                {
                    query = query.Where(c => c.A1Nome.Contains(Cliente) || c.A1Nreduz.Contains(Cliente));
                }

                ClientesDenuo = query.ToList();
            }

            return Page();
        }
    }
}
