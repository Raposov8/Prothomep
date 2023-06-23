using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Denuo;
using SGID.Models.Inter;
using SGID.Models.Patrimonio;

namespace SGID.Pages.Formularios.Patrimonio
{
    [Authorize]
    public class RelatorioPatrimonioIncompletoModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public DateTime Data { get;set; }
        public string Codigo { get; set; }
        public string Id { get; set; }

        public List<Models.Patrimonio.Patrimonio> Relatorio { get; set; } = new List<Models.Patrimonio.Patrimonio>();
        public RelatorioPatrimonioIncompletoModel(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
        }

        public void OnGet(string Id,string Empresa,string Filial)
        {
            Data = DateTime.Now;
            this.Id= Id;

            if (Empresa == "01")
            {
                Codigo = ProtheusInter.Pa1010s.FirstOrDefault(x => x.DELET != "*" && x.Pa1Codigo == Id && x.Pa1Filial == Filial).Pa1Despat;

                Relatorio = (from PA20 in ProtheusInter.Pa2010s
                             join SB10 in ProtheusInter.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                             where PA20.DELET != "*" && SB10.DELET != "*" && PA20.Pa2Codigo == Id && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                             orderby SB10.B1Cod
                             select new Models.Patrimonio.Patrimonio
                             {
                                 CodProd = SB10.B1Cod,
                                 DescProd = SB10.B1Desc,
                                 QtdKit = PA20.Pa2Qtdkit,
                                 QtdPat = PA20.Pa2Qtdpat
                             }).ToList();
            }
            else
            {
                var PAtri = ProtheusDenuo.Pa1010s.FirstOrDefault(x => x.DELET != "*" && x.Pa1Codigo == Id && x.Pa1Filial == Filial);
                Codigo = PAtri.Pa1Despat;

                Relatorio = (from PA20 in ProtheusDenuo.Pa2010s
                             join SB10 in ProtheusDenuo.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                             where PA20.DELET != "*" && SB10.DELET != "*" && PA20.Pa2Codigo == Id && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                             && PA20.Pa2Filial == PAtri.Pa1Filial
                             orderby SB10.B1Cod
                             select new Models.Patrimonio.Patrimonio
                             {
                                 CodProd = SB10.B1Cod,
                                 DescProd = SB10.B1Desc,
                                 QtdKit = PA20.Pa2Qtdkit,
                                 QtdPat = PA20.Pa2Qtdpat
                             }).ToList();
            }
        }
    }
}
