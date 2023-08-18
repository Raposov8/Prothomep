using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;

namespace SGID.Pages.Logistica
{
    public class LogisticaComercialModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<Agendamentos> Agendamentos { get; set; }
        public int EmTrasporte { get; set; }

        public int Entregue { get; set; }

        public LogisticaComercialModel(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter, ApplicationDbContext sgid)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
            SGID = sgid;
        }

        public void OnGet(int Id)
        {
            var user = User.Identity.Name.Split("@")[0].ToUpper();

            EmTrasporte = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.StatusLogistica == 1 && x.VendedorLogin == user)
                .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).Count();

            Entregue = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.StatusLogistica == 2 && x.VendedorLogin == user)
                .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).Count();

            if(Id == 1)
            {
                Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.StatusLogistica == 1 && x.VendedorLogin == user && x.DataCirurgia != null)
                .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList();
            }
            else
            {
                Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.StatusLogistica == 2 && x.VendedorLogin == user && x.DataCirurgia != null)
                .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList();
            }

        }

        
    }
}
