using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Models.Denuo;
using SGID.Models.Patrimonio;

namespace SGID.Pages.Formularios
{
    [Authorize]
    public class CONSPA3Model : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        public List<DISPPA3> Relatorio { get; set; }
        public string Patrimonio { get; set; }
        public string NumPatri { get; set; }
        public string Empresa { get; set; }
        public DateTime Data { get; set; }
        public CONSPA3Model(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
        }

        public void OnGet(string id)
        {
            Empresa = id;
        }

        public IActionResult OnPost(string Patrimonio, string Empresa)
        {
            this.Empresa = Empresa;
            Data = DateTime.Now;
            Relatorio = new List<DISPPA3>();
            if (Empresa == "01")
            {
                this.NumPatri = Patrimonio;
                Relatorio = (from PA30 in ProtheusInter.Pa3010s
                             join SB10 in ProtheusInter.Sb1010s on PA30.Pa3Produt equals SB10.B1Cod
                             where PA30.DELET != "*" && SB10.DELET != "*"
                             && PA30.Pa3Codigo == Patrimonio
                             orderby PA30.Pa3Produt, PA30.Pa3Lote
                             select new DISPPA3
                             {
                                 Codigo = PA30.Pa3Codigo,
                                 Produto = PA30.Pa3Produt,
                                 Descricao = SB10.B1Desc,
                                 QtdPat = PA30.Pa3Qtdpat,
                                 QtdLot = PA30.Pa3Qtdlot,
                                 Lote = PA30.Pa3Lote
                             }
                             ).ToList();
            }
            else
            {
                this.NumPatri = Patrimonio;
                Relatorio = (from PA30 in ProtheusDenuo.Pa3010s
                             join SB10 in ProtheusDenuo.Sb1010s on PA30.Pa3Produt equals SB10.B1Cod
                             where PA30.DELET != "*" && SB10.DELET != "*"
                             && PA30.Pa3Codigo == Patrimonio
                             orderby PA30.Pa3Produt,PA30.Pa3Lote
                             select new DISPPA3
                             {
                                 Codigo = PA30.Pa3Codigo,
                                 Produto = PA30.Pa3Produt,
                                 Descricao = SB10.B1Desc,
                                 QtdPat = PA30.Pa3Qtdpat,
                                 QtdLot = PA30.Pa3Qtdlot,
                                 Lote = PA30.Pa3Lote
                             }
                             ).ToList();
            }

            

            return Page();
        }
    }

}
