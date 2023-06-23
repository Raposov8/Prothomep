using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Estoque;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.Estoque
{

    [Authorize]
    public class PA3Model : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioPA3> Relatorio { get; set; } = new List<RelatorioPA3>();

        public PA3Model(TOTVSINTERContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public string Patrimonio { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string Patrimonio)
        {
            try
            {
                this.Patrimonio = Patrimonio;

                Relatorio = (from PA30 in Protheus.Pa3010s
                             join SB10 in Protheus.Sb1010s on PA30.Pa3Produt equals SB10.B1Cod
                             where PA30.DELET != "*" && SB10.DELET != "*" && PA30.Pa3Codigo == Patrimonio
                             select new RelatorioPA3
                             {
                                 Produto = PA30.Pa3Produt,
                                 Descricao = SB10.B1Desc,
                                 QtdePatrim = PA30.Pa3Qtdpat,
                                 QtdeLote = PA30.Pa3Qtdlot,
                                 Lote = PA30.Pa3Lote
                             }).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "PA3",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {

                Relatorio = (from PA30 in Protheus.Pa3010s
                             join SB10 in Protheus.Sb1010s on PA30.Pa3Produt equals SB10.B1Cod
                             where PA30.DELET != "*" && SB10.DELET != "*" && PA30.Pa3Codigo == Patrimonio
                             select new RelatorioPA3
                             {
                                 Produto = PA30.Pa3Produt,
                                 Descricao = SB10.B1Desc,
                                 QtdePatrim = PA30.Pa3Qtdpat,
                                 QtdeLote = PA30.Pa3Qtdlot,
                                 Lote = PA30.Pa3Lote
                             }).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("PA3");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "PA3");

                sheet.Cells[1, 1].Value = "Produto";
                sheet.Cells[1, 2].Value = "Descrição";
                sheet.Cells[1, 3].Value = "Qtde Patrim.";
                sheet.Cells[1, 4].Value = "Qtde Lote";
                sheet.Cells[1, 5].Value = "Lote";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Produto;
                    sheet.Cells[i, 2].Value = Pedido.Descricao;
                    sheet.Cells[i, 3].Value = Pedido.QtdePatrim;
                    sheet.Cells[i, 4].Value = Pedido.QtdePatrim;
                    sheet.Cells[i, 5].Value = Pedido.Lote;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PA3.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "PA3 Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
