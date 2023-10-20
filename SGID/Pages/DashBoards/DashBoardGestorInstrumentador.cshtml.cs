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

        public DashBoardGestorInstrumentadorModel(ApplicationDbContext sgid)
        {
            SGID = sgid;

        }
        public void OnGet()
        {

            Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 3).ToList();

        }
    }
}
