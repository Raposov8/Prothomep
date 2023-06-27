using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.Clientes
{
    public class MedicosAreaModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public List<Models.Denuo.Sa1010> ClientesDenuo { get; set; } = new List<Models.Denuo.Sa1010>();
        public List<Models.Inter.Sa1010> ClientesInter { get; set; } = new List<Models.Inter.Sa1010>();

        public string Bairro { get; set; } = "";

        public MedicosAreaModel(TOTVSDENUOContext denuo, TOTVSINTERContext inter)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
        }
        public void OnGet(string Bairro)
        {
            if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
            {
                ClientesInter = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Msblql != "1" && x.A1Bairro == Bairro).ToList();
            }
            else
            {
                ClientesDenuo = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Msblql != "1" && x.A1Bairro == Bairro).ToList();
            }

            this.Bairro = Bairro;
        }
    }
}
