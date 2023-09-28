using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;

namespace SGID.Pages.Instrumentador
{
    public class ListarInstrumentadoresModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        public List<RecursoIntrumentador> Instrumentadores { get; set; } = new List<RecursoIntrumentador>();

        public ListarInstrumentadoresModel(ApplicationDbContext sgid) => SGID = sgid;

        public void OnGet()
        {
            Instrumentadores = SGID.Instrumentadores.ToList();
        }

        
    }
}
