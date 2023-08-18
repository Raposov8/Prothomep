using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Models.Denuo;

namespace SGID.Pages.Formularios.Patrimonio
{
    [Authorize]
    public class RelatorioPatrimonioCompletoModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public DateTime Data { get; set; }
        public string Codigo { get; set; }
        public string Id { get; set; }

        public List<Models.Patrimonio.Patrimonio> Relatorio { get; set; } = new List<Models.Patrimonio.Patrimonio>();
        public RelatorioPatrimonioCompletoModel(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
        }

        public void OnGet(string Id,string Empresa,string Filial)
        {
            Data = DateTime.Now;
            this.Id = Id;

            if (Empresa == "01")
            {
                Codigo = ProtheusInter.Pa1010s.FirstOrDefault(x => x.DELET != "*" && x.Pa1Codigo == Id && x.Pa1Filial == Filial).Pa1Despat;

                Relatorio = (from PA20 in ProtheusInter.Pa2010s
                             join SB10 in ProtheusInter.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                             where PA20.DELET != "*" && SB10.DELET != "*" && PA20.Pa2Codigo == Id && PA20.Pa2Qtdkit != 0
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
                
                var PAtri = ProtheusDenuo.Pa1010s.FirstOrDefault(x => x.DELET != "*" && x.Pa1Codigo == Id && x.Pa1Filial == Filial && ((int)(object)x.Pa1Dtinsp >= 20200701 || x.Pa1Dtinsp == null));

                Codigo = PAtri.Pa1Despat;

                Relatorio = (from PA20 in ProtheusDenuo.Pa2010s
                             join SB10 in ProtheusDenuo.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                             where PA20.DELET != "*" && SB10.DELET != "*" && PA20.Pa2Codigo == Id && PA20.Pa2Qtdkit != 0
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
