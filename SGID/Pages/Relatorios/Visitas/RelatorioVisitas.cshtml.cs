using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Models.Denuo;
using SGID.Models.Diretoria;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.Visitas
{
    [Authorize]
    public class RelatorioVisitasModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        private TOTVSINTERContext INTER { get; set; }
        private TOTVSDENUOContext DENUO { get; set; }

        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public List<Data.ViewModel.Visitas> Visitas { get; set; } = new List<Data.ViewModel.Visitas>();
        public List<Data.ViewModel.Visitas> VisitasFilter { get; set; } = new List<Data.ViewModel.Visitas>();
        public List<RankingVendedores> Vendedores { get; set; } = new List<RankingVendedores>();
        public string Linhas { get; set; }

        public RelatorioVisitasModel(ApplicationDbContext sgid,TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            SGID = sgid;
            INTER = inter;
            DENUO = denuo;
        }

        public void OnGet(string Linha)
        {
            Linhas = Linha;
        }


        public IActionResult OnPost(DateTime DataInicio, DateTime DataFim,string Linha)
        {

            Inicio = DataInicio;

            Fim = DataFim;

            var VendedoresInter = INTER.Sa3010s.Where(x => x.DELET != "*" && x.A3Msblql != "1").ToList();

            var VendedoresDenuo = DENUO.Sa3010s.Where(x => x.DELET != "*" && x.A3Msblql != "1").ToList();

            Visitas = SGID.Visitas.Where(x => x.DataHora >= Inicio && x.DataHora <= Fim).ToList();

            Vendedores = SGID.Visitas.GroupBy(x => x.Vendedor).Select(x => new RankingVendedores { Nome = x.Key }).ToList();

            Vendedores.ForEach(x =>
            {
                x.Linha = VendedoresInter.FirstOrDefault(c => c.A3Xlogin.Trim() == x.Nome)?.A3Xdescun.Trim();

                x.Linha ??= VendedoresDenuo.FirstOrDefault(c => c.A3Xlogin.Trim() == x.Nome)?.A3Xdescun.Trim();
            });

            Vendedores = Vendedores.Where(x => x.Linha == Linha).ToList();

            Vendedores.ForEach(vendedor =>
            {
                var vistas = Visitas.Where(x => x.Vendedor == vendedor.Nome).ToList();

                vendedor.Quant = vistas.Count;
            });

            Vendedores = Vendedores.OrderByDescending(x => x.Quant).ToList();

            Vendedores.ForEach(x =>
            {
                VisitasFilter.AddRange(Visitas.Where(c => c.Vendedor == x.Nome).ToList());
            });

            VisitasFilter = VisitasFilter.OrderBy(x => x.DataHora).ToList();

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
