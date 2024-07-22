using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using SGID.Models.Denuo;
using SGID.Models.Inter;
using SGID.Models.Patrimonio;
using SGID.Data.Models;
using SGID.Data;

namespace SGID.Pages.Formularios.Patrimonio
{
    [Authorize]
    public class RelatorioTodosItensPatrimoniosModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        private ApplicationDbContext SGID { get; set; }

        public DateTime Data { get; set; }
        public string Codigo { get; set; }
        public string Id { get; set; }

        public List<Models.Patrimonio.Patrimonio> Relatorio { get; set; } = new List<Models.Patrimonio.Patrimonio>();
        public RelatorioTodosItensPatrimoniosModel(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter, ApplicationDbContext sgid)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
            SGID = sgid;
        }

        public void OnGet(string Empresa, string Filial)
        {
            Data = DateTime.Now;
            if (Empresa == "01")
            {
                Relatorio = (from PA20 in ProtheusInter.Pa2010s
                             join PA10 in ProtheusInter.Pa1010s on new { Codigo = PA20.Pa2Codigo, Filial = PA20.Pa2Filial } equals new { Codigo = PA10.Pa1Codigo, Filial = PA10.Pa1Filial }
                             join SB10 in ProtheusInter.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                             where PA20.DELET != "*" && SB10.DELET != "*" && PA20.Pa2Qtdkit != 0
                             && PA10.DELET != "*"
                             orderby SB10.B1Cod
                             select new Models.Patrimonio.Patrimonio
                             {
                                 CodPatri = PA20.Pa2Codigo,
                                 DescPatri = PA10.Pa1Despat,
                                 CodProd = SB10.B1Cod,
                                 DescProd = SB10.B1Desc,
                                 QtdKit = PA20.Pa2Qtdkit,
                                 QtdPat = PA20.Pa2Qtdpat,
                                 Bloqueio = PA10.Pa1Msblql
                             }).OrderBy(x => x.CodPatri).ToList();
            }
            else
            {

                Relatorio = (from PA20 in ProtheusDenuo.Pa2010s
                             join PA10 in ProtheusDenuo.Pa1010s on new { Codigo = PA20.Pa2Codigo, Filial = PA20.Pa2Filial } equals new { Codigo = PA10.Pa1Codigo, Filial = PA10.Pa1Filial }
                             join SB10 in ProtheusDenuo.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                             where PA20.DELET != "*" && SB10.DELET != "*" && PA20.Pa2Qtdkit != 0
                             && ((int)(object)PA10.Pa1Dtinsp >= 20200701 || PA10.Pa1Dtinsp == null)
                             && PA20.Pa2Filial == "03" && PA10.DELET != "*"
                             orderby SB10.B1Cod
                             select new Models.Patrimonio.Patrimonio
                             {
                                 CodPatri = PA20.Pa2Codigo,
                                 DescPatri = PA10.Pa1Despat,
                                 CodProd = SB10.B1Cod,
                                 DescProd = SB10.B1Desc,
                                 QtdKit = PA20.Pa2Qtdkit,
                                 QtdPat = PA20.Pa2Qtdpat,
                                 Bloqueio = PA10.Pa1Msblql
                             }).OrderBy(x => x.CodPatri).ToList();
            }
        }


        public IActionResult OnPostExport(string Empresa, string Filial)
        {
            try
            {
                Data = DateTime.Now;
                if (Empresa == "01")
                {
                    Relatorio = (from PA20 in ProtheusInter.Pa2010s
                                 join PA10 in ProtheusInter.Pa1010s on new { Codigo = PA20.Pa2Codigo, Filial = PA20.Pa2Filial } equals new { Codigo = PA10.Pa1Codigo, Filial = PA10.Pa1Filial }
                                 join SB10 in ProtheusInter.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                                 where PA20.DELET != "*" && SB10.DELET != "*" && PA20.Pa2Qtdkit != 0
                                 && PA10.DELET != "*"
                                 orderby SB10.B1Cod
                                 select new Models.Patrimonio.Patrimonio
                                 {
                                     CodPatri = PA20.Pa2Codigo,
                                     DescPatri = PA10.Pa1Despat,
                                     CodProd = SB10.B1Cod,
                                     DescProd = SB10.B1Desc,
                                     QtdKit = PA20.Pa2Qtdkit,
                                     QtdPat = PA20.Pa2Qtdpat,
                                     Bloqueio = PA10.Pa1Msblql
                                 }).OrderBy(x => x.CodPatri).ToList();
                }
                else
                {

                    Relatorio = (from PA20 in ProtheusDenuo.Pa2010s
                                 join PA10 in ProtheusDenuo.Pa1010s on new { Codigo = PA20.Pa2Codigo, Filial = PA20.Pa2Filial } equals new { Codigo = PA10.Pa1Codigo, Filial = PA10.Pa1Filial }
                                 join SB10 in ProtheusDenuo.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                                 where PA20.DELET != "*" && SB10.DELET != "*" && PA20.Pa2Qtdkit != 0
                                 && ((int)(object)PA10.Pa1Dtinsp >= 20200701 || PA10.Pa1Dtinsp == null)
                                 && PA20.Pa2Filial == "03" && PA10.DELET != "*"
                                 orderby SB10.B1Cod
                                 select new Models.Patrimonio.Patrimonio
                                 {
                                     CodPatri = PA20.Pa2Codigo,
                                     DescPatri = PA10.Pa1Despat,
                                     CodProd = SB10.B1Cod,
                                     DescProd = SB10.B1Desc,
                                     QtdKit = PA20.Pa2Qtdkit,
                                     QtdPat = PA20.Pa2Qtdpat,
                                     Bloqueio = PA10.Pa1Msblql
                                 }).OrderBy(x=> x.CodPatri).ToList();
                }

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("RelatorioTodosItensPatrimonios");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "RelatorioTodosItensPatrimonios");

                sheet.Cells[1, 1].Value = "NUM PATRIMONIO(S):";
                sheet.Cells[1, 2].Value = "DESCRIÇÃO PATRIMONIO(S):";
                sheet.Cells[1, 3].Value = "CODIGO(S):";
                sheet.Cells[1, 4].Value = "DESCRIÇÃO:";
                sheet.Cells[1, 5].Value = "QUANTIDADE / LOTE VIRTUAL:";
                sheet.Cells[1, 6].Value = "QUANTIDADE / LOTE FISICO";
                sheet.Cells[1, 7].Value = "BLOQUEIO";
                sheet.Cells[1, 8].Value = "OBSERVAÇÃO";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.CodPatri;
                    sheet.Cells[i, 2].Value = Pedido.DescPatri;
                    sheet.Cells[i, 3].Value = Pedido.CodProd;
                    sheet.Cells[i, 4].Value = Pedido.DescProd;
                    sheet.Cells[i, 5].Value = Pedido.QtdKit;
                    sheet.Cells[i, 6].Value = Pedido.QtdPat;
                    sheet.Cells[i, 7].Value = Pedido.Bloqueio;

                    i++;
                });


                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DISPPATRT.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "DISPPATRT Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
