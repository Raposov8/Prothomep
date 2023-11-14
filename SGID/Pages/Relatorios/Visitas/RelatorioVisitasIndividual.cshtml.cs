using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Models.Diretoria;
using SGID.Models.Visitas;

namespace SGID.Pages.Relatorios.Visitas
{
    [Authorize]
    public class RelatorioVisitasIndividualModel : PageModel
    {

        private ApplicationDbContext SGID { get; set; }

        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public List<UltimasVisitas> RankingUltimas { get; set; } = new List<UltimasVisitas>();

        public List<Data.ViewModel.Visitas> Visitas { get; set; } = new List<Data.ViewModel.Visitas>();

        public List<Bairros> MaxBairros { get; set; } = new List<Bairros>(); 

        public RelatorioVisitasIndividualModel(ApplicationDbContext sgid)
        {
            SGID = sgid;
        }

        public IActionResult OnGet(string Vendedor, DateTime DataInicio, DateTime DataFim)
        {

            Inicio = DataInicio;

            Fim = DataFim;

            Visitas = SGID.Visitas.Where(x => x.DataHora >= Inicio && x.DataHora <= Fim && x.Vendedor == Vendedor).OrderBy(x=> x.DataHora).ToList();

            var bairros = SGID.Visitas.Where(x => x.Vendedor == Vendedor && x.Status != 3).ToList();

            MaxBairros = bairros.GroupBy(x => x.Bairro)
                .Select(x => new Bairros { Nome = x.Key,Numero = x.Count()}).OrderByDescending(x=> x.Numero).Take(3).ToList();

            var Ultimas = bairros.GroupBy(x=> x.Medico)
                            .Select(x=> new UltimasVisitas
                            {
                                Nome = x.Key,
                                Data = x.OrderByDescending(x=> x.DataHora).First().DataHora
                            }).ToList();

            RankingUltimas = Ultimas.OrderBy(x => x.Data).Take(3).ToList();


            return Page();
        }
    }
}
