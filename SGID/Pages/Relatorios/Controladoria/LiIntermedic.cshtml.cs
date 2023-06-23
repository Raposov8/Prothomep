using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Inter;
using SGID.Models.Relatorio;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class LiIntermedicModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioLiIntermedic> Relatorio = new List<RelatorioLiIntermedic>();
        public LiIntermedicModel(TOTVSINTERContext context, ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        
        public string NLi { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string NumLi)
        {
            try
            {
                NLi = NumLi;

                Relatorio = (from SW50 in Protheus.Sw5010s
                             join SWP10 in Protheus.Swp010s on new { Num = SW50.W5PgiNum, Seq = SW50.W5SeqLi } equals new { Num = SWP10.WpPgiNum, Seq = SWP10.WpSeqLi }
                             join SB10 in Protheus.Sb1010s on SW50.W5CodI equals SB10.B1Cod
                             where SW50.DELET != "*" && SWP10.DELET != "*" && SB10.DELET != "*" &&
                             SW50.W5PgiNum == NumLi
                             select new RelatorioLiIntermedic
                             {
                                 NumLi = SW50.W5PgiNum,
                                 SeqLi = SW50.W5SeqLi,
                                 CodProd = SW50.W5CodI,
                                 Produto = SB10.B1Desc,
                                 QTDE = SW50.W5Qtde,
                                 SaldoQ = SW50.W5SaldoQ,
                                 ValUni = SW50.W5Preco,
                                 Total = SW50.W5Qtde * SW50.W5Preco,
                                 Ncm = SWP10.WpNcm,
                                 RegAnvisa = SB10.B1Reganvi,
                                 Validade = SB10.B1Vigente.ToUpper() == "VIGENTE" ? SB10.B1Vigente : $"{SB10.B1Dtvalid.Substring(6, 2)}/{SB10.B1Dtvalid.Substring(4, 2)}/{SB10.B1Dtvalid.Substring(0, 4)}",
                             }
                              ).ToList();

                Relatorio = Relatorio.OrderByDescending(x => x.Validade).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "LiIntermedic",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {

                Relatorio = (from SW50 in Protheus.Sw5010s
                             join SWP10 in Protheus.Swp010s on new { Num = SW50.W5PgiNum, Seq = SW50.W5SeqLi } equals new { Num = SWP10.WpPgiNum, Seq = SWP10.WpSeqLi }
                             join SB10 in Protheus.Sb1010s on SW50.W5CodI equals SB10.B1Cod
                             where SW50.DELET != "*" && SWP10.DELET != "*" && SB10.DELET != "*" &&
                             SW50.W5PgiNum == NLi
                             select new RelatorioLiIntermedic
                             {
                                 NumLi = SW50.W5PgiNum,
                                 SeqLi = SW50.W5SeqLi,
                                 CodProd = SW50.W5CodI,
                                 Produto = SB10.B1Desc,
                                 QTDE = SW50.W5Qtde,
                                 SaldoQ = SW50.W5SaldoQ,
                                 ValUni = SW50.W5Preco,
                                 Total = SW50.W5Qtde * SW50.W5Preco,
                                 Ncm = SWP10.WpNcm,
                                 RegAnvisa = SB10.B1Reganvi,
                                 Validade = SB10.B1Vigente.ToUpper() == "VIGENTE" ? SB10.B1Vigente : $"{SB10.B1Dtvalid.Substring(6, 2)}/{SB10.B1Dtvalid.Substring(4, 2)}/{SB10.B1Dtvalid.Substring(0, 4)}",
                             }
                              ).ToList();

                Relatorio = Relatorio.OrderByDescending(x => x.Validade).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("LiIntermedic");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "LiIntermedic");

                sheet.Cells[1, 1].Value = "NUM LI";
                sheet.Cells[1, 2].Value = "SEQ.LI";
                sheet.Cells[1, 3].Value = "COD PROD.";
                sheet.Cells[1, 4].Value = "PRODUTO";
                sheet.Cells[1, 5].Value = "QTDE";
                sheet.Cells[1, 6].Value = "SALDO Q";
                sheet.Cells[1, 7].Value = "VL UNIT";
                sheet.Cells[1, 8].Value = "TOTAL";
                sheet.Cells[1, 9].Value = "NCM";
                sheet.Cells[1, 10].Value = "REG.ANVISA";
                sheet.Cells[1, 11].Value = "VALIDADE";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.NumLi;
                    sheet.Cells[i, 2].Value = Pedido.SeqLi;
                    sheet.Cells[i, 3].Value = Pedido.CodProd;
                    sheet.Cells[i, 4].Value = Pedido.Produto;
                    sheet.Cells[i, 5].Value = Pedido.QTDE;
                    sheet.Cells[i, 6].Value = Pedido.SaldoQ;
                    sheet.Cells[i, 7].Value = Pedido.ValUni;
                    sheet.Cells[i, 8].Value = Pedido.Total;
                    sheet.Cells[i, 9].Value = Pedido.Ncm;
                    sheet.Cells[i, 10].Value = Pedido.RegAnvisa;
                    sheet.Cells[i, 11].Value = Pedido.Validade;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "LiIntermedic.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "LiIntermedic Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
