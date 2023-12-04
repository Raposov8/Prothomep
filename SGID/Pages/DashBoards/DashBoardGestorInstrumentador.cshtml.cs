using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;

namespace SGID.Pages.DashBoards
{
    public class DashBoardGestorInstrumentadorModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        public List<Agendamentos> Agendamentos { get; set; } = new List<Agendamentos>();

        public int Agendas { get; set; }
        public int Cirurgias { get; set; }
        public int Andamentos { get; set; }
        public int Canceladas { get; set; }

        public DashBoardGestorInstrumentadorModel(ApplicationDbContext sgid)
        {
            SGID = sgid;

        }
        public void OnGet()
        {


            Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 7).ToList();



        }
    }
}
