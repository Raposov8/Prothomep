using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Models.Denuo;

namespace SGID.Pages.Formularios
{
    [Authorize]
    public class REQ30Model : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        public List<Models.Patrimonio.Patrimonio> Relatorio { get; set; }
        public string Patrimonio { get; set; }
        public string NumPatri { get; set; }
        public string NomPatri { get; set; }
        public string Empresa { get; set; }
        public DateTime Data { get; set; }
        public REQ30Model(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
        }

        public void OnGet(string id)
        {
            this.Empresa = id;
        }

        public IActionResult OnPost(string Patrimonio, string Empresa,string Filial)
        {
            this.Empresa = Empresa;
            Data = DateTime.Now;
            Relatorio = new List<Models.Patrimonio.Patrimonio>();
            if (Empresa == "01")
            {
                NumPatri = Patrimonio;
                Relatorio = (from PA20 in ProtheusInter.Pa2010s
                             join SB10 in ProtheusInter.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                             where PA20.DELET != "*" && SB10.DELET != "*"
                             && PA20.Pa2Codigo == Patrimonio  && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                             select new Models.Patrimonio.Patrimonio
                             {
                                 CodProd = SB10.B1Cod,
                                 DescProd = SB10.B1Desc,
                                 QtdKit = PA20.Pa2Qtdkit,
                                 QtdPat = PA20.Pa2Qtdpat
                             }
                             ).ToList();

                NomPatri = ProtheusInter.Pa1010s.FirstOrDefault(x => x.DELET != "*" && x.Pa1Codigo == Patrimonio)?.Pa1Despat;
            }
            else
            {
                NumPatri = Patrimonio;
                Relatorio = (from PA20 in ProtheusDenuo.Pa2010s
                             join SB10 in ProtheusDenuo.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                             where PA20.DELET != "*" && SB10.DELET != "*"
                             && PA20.Pa2Codigo == Patrimonio && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                             && PA20.Pa2Filial == Filial
                             select new Models.Patrimonio.Patrimonio
                             {
                                 CodProd = SB10.B1Cod,
                                 DescProd = SB10.B1Desc,
                                 QtdKit = PA20.Pa2Qtdkit,
                                 QtdPat = PA20.Pa2Qtdpat
                             }
                             ).ToList();
                NomPatri = ProtheusDenuo.Pa1010s.FirstOrDefault(x => x.DELET != "*" && x.Pa1Codigo == Patrimonio && x.Pa1Filial == Filial)?.Pa1Despat;
            }

            return Page();
        }
    }
}
