using DocumentFormat.OpenXml.Office2010.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using OPMEnexo;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Estoque;
using SGID.Models.Estoque.RelatorioFaturamentoNFFab;
using SGID.Models.Denuo;

namespace SGID.Pages.Relatorios.Estoque
{

    [Authorize]
    public class SaldoEstoqueModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioSaldoEstoque> Relatorio { get; set; } = new List<RelatorioSaldoEstoque>();
        public List<string> Locals { get; set; } = new List<string>();

        public SaldoEstoqueModel(TOTVSDENUOContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }

        public void OnGet()
        {
            try
            {

                Relatorio = (from SB80 in Protheus.Sb8010s
                             join SB10 in Protheus.Sb1010s on SB80.B8Produto equals SB10.B1Cod
                             where SB10.DELET != "*" && SB80.DELET != "*" && SB80.B8Saldo != 0
                             select new RelatorioSaldoEstoque
                             {
                                 Filial = SB80.B8Filial,
                                 Produto = SB80.B8Produto,
                                 DescProd = SB10.B1Desc,
                                 LoteCtl = SB80.B8Lotectl,
                                 Local = SB80.B8Local,
                                 Saldo = SB80.B8Saldo,
                                 UM = SB10.B1Um,
                                 Empenho = SB80.B8Empenho,
                                 Data = $"{SB80.B8Data.Substring(6, 2)}/{SB80.B8Data.Substring(4, 2)}/{SB80.B8Data.Substring(0, 4)}",
                                 DtValid = $"{SB80.B8Dtvalid.Substring(6, 2)}/{SB80.B8Dtvalid.Substring(4, 2)}/{SB80.B8Dtvalid.Substring(0, 4)}",
                                 Descendereco = ""
                             }).ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SaldoEstoque",user);
            }
        }

        public IActionResult OnPost(string Local)
        {
            try
            {

                Relatorio = (from SB80 in Protheus.Sb8010s
                             join SB10 in Protheus.Sb1010s on SB80.B8Produto equals SB10.B1Cod
                             where SB10.DELET != "*" && SB80.DELET != "*" && SB80.B8Saldo != 0
                             && SB80.B8Local==Local
                             select new RelatorioSaldoEstoque
                             {
                                 Filial = SB80.B8Filial,
                                 Produto = SB80.B8Produto,
                                 DescProd = SB10.B1Desc,
                                 LoteCtl = SB80.B8Lotectl,
                                 Local = SB80.B8Local,
                                 Saldo = SB80.B8Saldo,
                                 UM = SB10.B1Um,
                                 Empenho = SB80.B8Empenho,
                                 Data = $"{SB80.B8Data.Substring(6, 2)}/{SB80.B8Data.Substring(4, 2)}/{SB80.B8Data.Substring(0, 4)}",
                                 DtValid = $"{SB80.B8Dtvalid.Substring(6, 2)}/{SB80.B8Dtvalid.Substring(4, 2)}/{SB80.B8Dtvalid.Substring(0, 4)}",
                                 Descendereco = ""
                             }).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SaldoEstoque", user);
                return Page();
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {
                Relatorio = (from SB80 in Protheus.Sb8010s
                             join SB10 in Protheus.Sb1010s on SB80.B8Produto equals SB10.B1Cod
                             where SB10.DELET != "*" && SB80.DELET != "*" && SB80.B8Saldo != 0
                             select new RelatorioSaldoEstoque
                             {
                                 Filial = SB80.B8Filial,
                                 Produto = SB80.B8Produto,
                                 DescProd = SB10.B1Desc,
                                 LoteCtl = SB80.B8Lotectl,
                                 Local = SB80.B8Local,
                                 Endereco = "",
                                 Saldo = SB80.B8Saldo,
                                 UM = SB10.B1Um,
                                 Empenho = 0,
                                 Data = $"{SB80.B8Data.Substring(6, 2)}/{SB80.B8Data.Substring(4, 2)}/{SB80.B8Data.Substring(0, 4)}",
                                 DtValid = $"{SB80.B8Dtvalid.Substring(6, 2)}/{SB80.B8Dtvalid.Substring(4, 2)}/{SB80.B8Dtvalid.Substring(0, 4)}",
                                 Descendereco = ""
                             }).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("SaldoEstoque");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "SaldoEstoque");

                sheet.Cells[1, 1].Value = "Filial";
                sheet.Cells[1, 2].Value = "Produto";
                sheet.Cells[1, 3].Value = "Desc Prod.";
                sheet.Cells[1, 4].Value = "Lotectl";
                sheet.Cells[1, 5].Value = "Local";
                sheet.Cells[1, 6].Value = "Endereço";
                sheet.Cells[1, 7].Value = "Saldo";
                sheet.Cells[1, 8].Value = "UM";
                sheet.Cells[1, 9].Value = "Empenho";
                sheet.Cells[1, 10].Value = "Data";
                sheet.Cells[1, 11].Value = "DtValid";
                sheet.Cells[1, 12].Value = "Descendereco";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.Produto;
                    sheet.Cells[i, 3].Value = Pedido.DescProd;
                    sheet.Cells[i, 4].Value = Pedido.LoteCtl;
                    sheet.Cells[i, 5].Value = Pedido.Local;
                    sheet.Cells[i, 6].Value = Pedido.Endereco;
                    sheet.Cells[i, 7].Value = Pedido.Saldo;
                    sheet.Cells[i, 8].Value = Pedido.UM;
                    sheet.Cells[i, 9].Value = Pedido.Empenho;
                    sheet.Cells[i, 10].Value = Pedido.Data;
                    sheet.Cells[i, 11].Value = Pedido.DtValid;
                    sheet.Cells[i, 12].Value = Pedido.Descendereco;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SaldoEstoque.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SaldoEstoque Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
