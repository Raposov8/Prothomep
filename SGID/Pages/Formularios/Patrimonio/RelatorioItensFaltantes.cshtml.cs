using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;

namespace SGID.Pages.Formularios.Patrimonio
{
    [Authorize]
    public class RelatorioItensFaltantesModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        private ApplicationDbContext SGID { get; set; }

        public DateTime Data { get; set; }
        public string Empresa { get; set; }

        public List<Models.Patrimonio.Patrimonio> Relatorio { get; set; } = new List<Models.Patrimonio.Patrimonio>();
        public RelatorioItensFaltantesModel(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter,ApplicationDbContext sgid)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
            SGID = sgid;
        }

        public void OnGet(string Empresa)
        {
            Data = DateTime.Now;

            this.Empresa = Empresa;

            if (Empresa == "01")
            {
                Relatorio = (from PA20 in ProtheusInter.Pa2010s
                             join SB10 in ProtheusInter.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                             join PA10 in ProtheusInter.Pa1010s on new {Codigo=PA20.Pa2Codigo,Filial=PA20.Pa2Filial } equals new { Codigo = PA10.Pa1Codigo, Filial = PA10.Pa1Filial }
                             where PA20.DELET != "*" && SB10.DELET != "*"  && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                             && PA10.DELET!="*" && PA10.Pa1Msblql != "1"
                             orderby PA20.Pa2Codigo
                             select new Models.Patrimonio.Patrimonio
                             {
                                 CodProd = SB10.B1Cod,
                                 CodPatri = PA20.Pa2Codigo,
                                 QuantFalt = PA20.Pa2Qtdkit - PA20.Pa2Qtdpat
                             }).ToList();
            }
            else
            {

                Relatorio = (from PA20 in ProtheusDenuo.Pa2010s
                             join SB10 in ProtheusDenuo.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                             join PA10 in ProtheusDenuo.Pa1010s on new { Codigo = PA20.Pa2Codigo, Filial = PA20.Pa2Filial } equals new { Codigo = PA10.Pa1Codigo, Filial = PA10.Pa1Filial }
                             where PA20.DELET != "*" && SB10.DELET != "*"  && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                             && PA10.DELET != "*" && PA10.Pa1Msblql != "1"
                             && PA10.Pa1Filial == "03"
                             orderby PA20.Pa2Codigo
                             select new Models.Patrimonio.Patrimonio
                             {
                                 CodProd = SB10.B1Cod,
                                 CodPatri = PA20.Pa2Codigo,
                                 QuantFalt = PA20.Pa2Qtdkit - PA20.Pa2Qtdpat
                             }).ToList();
            }
        }

        public IActionResult OnPost(string Empresa)
        {
            Data = DateTime.Now;

            this.Empresa = Empresa;
            var Nome = "";
            try
            {
                if (Empresa == "01")
                {
                    Relatorio = (from PA20 in ProtheusInter.Pa2010s
                                 join SB10 in ProtheusInter.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                                 join PA10 in ProtheusInter.Pa1010s on new { Codigo = PA20.Pa2Codigo, Filial = PA20.Pa2Filial } equals new { Codigo = PA10.Pa1Codigo, Filial = PA10.Pa1Filial }
                                 where PA20.DELET != "*" && SB10.DELET != "*" && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                                 && PA10.DELET != "*" && PA10.Pa1Msblql != "1"
                                 orderby PA20.Pa2Codigo
                                 select new Models.Patrimonio.Patrimonio
                                 {
                                     CodProd = SB10.B1Cod,
                                     CodPatri = PA20.Pa2Codigo,
                                     QuantFalt = PA20.Pa2Qtdkit - PA20.Pa2Qtdpat
                                 }).ToList();

                    Nome = "PatrimoniosItensFalantesInter";
                }
                else
                {

                    Relatorio = (from PA20 in ProtheusDenuo.Pa2010s
                                 join SB10 in ProtheusDenuo.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                                 join PA10 in ProtheusDenuo.Pa1010s on new { Codigo = PA20.Pa2Codigo, Filial = PA20.Pa2Filial } equals new { Codigo = PA10.Pa1Codigo, Filial = PA10.Pa1Filial }
                                 where PA20.DELET != "*" && SB10.DELET != "*" && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                                 && PA10.DELET != "*" && PA10.Pa1Msblql != "1"
                                 && PA10.Pa1Filial == "03"
                                 orderby PA20.Pa2Codigo
                                 select new Models.Patrimonio.Patrimonio
                                 {
                                     CodProd = SB10.B1Cod,
                                     CodPatri = PA20.Pa2Codigo,
                                     QuantFalt = PA20.Pa2Qtdkit - PA20.Pa2Qtdpat
                                 }).ToList();

                    Nome = "PatrimoniosItensFalantesDenuo";
                }


                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Itens Faltantes");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Itens Faltantes");

                sheet.Cells[1, 1].Value = "NUMERO";
                sheet.Cells[1, 2].Value = "CODIGO";
                sheet.Cells[1, 3].Value = "QUANT. FALTANTE";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.CodPatri;
                    sheet.Cells[i, 2].Value = Pedido.CodProd;
                    sheet.Cells[i, 3].Value = Pedido.QuantFalt;

                    i++;
                });

                

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{Nome}.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "PatrimoniosItensFalantes Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
