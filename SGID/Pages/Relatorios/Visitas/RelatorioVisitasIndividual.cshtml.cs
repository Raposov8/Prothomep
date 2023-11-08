using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Models.Diretoria;

namespace SGID.Pages.Relatorios.Visitas
{
    [Authorize]
    public class RelatorioVisitasIndividualModel : PageModel
    {

        private ApplicationDbContext SGID { get; set; }

        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }


        public List<Data.ViewModel.Visitas> Visitas { get; set; } = new List<Data.ViewModel.Visitas>();

        public RelatorioVisitasIndividualModel(ApplicationDbContext sgid)
        {
            SGID = sgid;
        }

        public IActionResult OnGet(string Vendedor, DateTime DataInicio, DateTime DataFim)
        {

            Inicio = DataInicio;

            Fim = DataFim;

            Visitas = SGID.Visitas.Where(x => x.DataHora >= Inicio && x.DataHora <= Fim && x.Vendedor == Vendedor).OrderBy(x=> x.DataHora).ToList();

            return Page();
        }
    }
}
