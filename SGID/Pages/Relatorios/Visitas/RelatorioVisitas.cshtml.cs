using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Diretoria;

namespace SGID.Pages.Relatorios.Visitas
{
    public class RelatorioVisitasModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }


        public List<Data.ViewModel.Visitas> Visitas { get; set; } = new List<Data.ViewModel.Visitas>();
        public List<RankingVendedores> Vendedores { get; set; } = new List<RankingVendedores>();

        public RelatorioVisitasModel(ApplicationDbContext sgid)
        {
            SGID = sgid;
        }

        public void OnGet()
        {

        }


        public IActionResult OnPost(DateTime DataInicio, DateTime DataFim)
        {



            Inicio = DataInicio;

            Fim = DataFim;

            Visitas = SGID.Visitas.Where(x => x.DataHora >= Inicio && x.DataHora <= Fim).ToList();


            Vendedores = Visitas.GroupBy(x => x.Vendedor).Select(x => new RankingVendedores { Nome = x.Key }).ToList();


            Vendedores.ForEach(vendedor =>
            {

                var vistas = Visitas.Where(x => x.Vendedor == vendedor.Nome).ToList();


                vendedor.Quant = vistas.Count;
            });


            Vendedores = Vendedores.OrderByDescending(x => x.Quant).ToList();

            return Page();

        }

        /*public void OnGet(string Ano, string Mes)
        {
            var DataInicio = Convert.ToDateTime($"{Mes}/01/{Ano}");

            var DataFim = Convert.ToDateTime($"{Mes}/31/{Ano}");

            Visitas = SGID.Visitas.Where(x => x.DataCriacao >= DataInicio && x.DataCriacao <= DataFim).ToList();


            Vendedores = Visitas.GroupBy(x=> x.Vendedor).Select(x => new RankingVendedores { Nome = x.Key }).ToList();


            Vendedores.ForEach(vendedor =>
            {

                var vistas = Visitas.Where(x => x.Vendedor == vendedor.Nome).ToList();


                vendedor.Quant = vistas.Count;
            });
            
        }*/
    }
}
