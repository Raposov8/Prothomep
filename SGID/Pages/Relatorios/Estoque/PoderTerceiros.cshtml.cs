using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Estoque;

namespace SGID.Pages.Relatorios.Estoque
{

    [Authorize]
    public class PoderTerceirosModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioPoderTerceiros> Relatorio { get; set; } = new List<RelatorioPoderTerceiros>();

        public PoderTerceirosModel(TOTVSINTERContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public string Lote { get; set; }
        public string Cliente { get; set; }
        public string Produto { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string Produto,string Cliente,string Lote)
        {
            try
            {
                this.Lote = Lote;
                this.Cliente = Cliente;
                this.Produto = Produto;

                Relatorio = (from SB60 in Protheus.Sb6010s
                             join SD20 in Protheus.Sd2010s on SB60.B6Ident equals SD20.D2Identb6
                             where SB60.DELET != "*" && SD20.DELET != "*" && SB60.B6Saldo != 0
                             && SB60.B6Produto == Produto && SB60.B6Clifor == Cliente && SD20.D2Lotectl == Lote
                             select new RelatorioPoderTerceiros
                             {
                                 Produto = SB60.B6Produto,
                                 Serie = SB60.B6Serie,
                                 NF = SB60.B6Doc,
                                 Emissao = $"{SB60.B6Emissao.Substring(6, 2)}/{SB60.B6Emissao.Substring(4, 2)}/{SB60.B6Emissao.Substring(0, 4)}",
                                 Tipo = "PODER3",
                                 Lote = SD20.D2Lotectl,
                                 Saldo = SB60.B6Saldo
                             }
                             )
                             .Concat
                             (from SD20 in Protheus.Sd2010s
                              join SF40 in Protheus.Sf4010s on SD20.D2Tes equals SF40.F4Codigo
                              where SD20.DELET != "*" && SF40.DELET != "*" && SF40.F4Poder3 == "N"
                              && SD20.D2Quant - SD20.D2Qtdedev != 0 && SD20.D2Cliente == Cliente
                              && SD20.D2Cod == Produto && SD20.D2Lotectl == Lote
                              select new RelatorioPoderTerceiros
                              {
                                  Produto = SD20.D2Cod,
                                  Serie = SD20.D2Serie,
                                  NF = SD20.D2Doc,
                                  Emissao = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                                  Tipo = "REMESSA",
                                  Lote = SD20.D2Lotectl,
                                  Saldo = SD20.D2Quant - SD20.D2Qtdedev
                              }
                             ).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "PoderTerceiros",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(string Produto, string Cliente, string Lote)
        {
            try
            {

                Relatorio = (from SB60 in Protheus.Sb6010s
                             join SD20 in Protheus.Sd2010s on SB60.B6Ident equals SD20.D2Identb6
                             where SB60.DELET != "*" && SD20.DELET != "*" && SB60.B6Saldo != 0
                             && SB60.B6Produto == Produto && SB60.B6Clifor == Cliente && SD20.D2Lotectl == Lote
                             select new RelatorioPoderTerceiros
                             {
                                 Produto = SB60.B6Produto,
                                 Serie = SB60.B6Serie,
                                 NF = SB60.B6Doc,
                                 Emissao = $"{SB60.B6Emissao.Substring(6, 2)}/{SB60.B6Emissao.Substring(4, 2)}/{SB60.B6Emissao.Substring(0, 4)}",
                                 Tipo = "PODER3",
                                 Lote = SD20.D2Lotectl,
                                 Saldo = SB60.B6Saldo
                             }
                             )
                             .Concat
                             (from SD20 in Protheus.Sd2010s
                              join SF40 in Protheus.Sf4010s on SD20.D2Tes equals SF40.F4Codigo
                              where SD20.DELET != "*" && SF40.DELET != "*" && SF40.F4Poder3 == "N"
                              && SD20.D2Quant - SD20.D2Qtdedev != 0 && SD20.D2Cliente == Cliente
                              && SD20.D2Cod == Produto && SD20.D2Lotectl == Lote
                              select new RelatorioPoderTerceiros
                              {
                                  Produto = SD20.D2Cod,
                                  Serie = SD20.D2Serie,
                                  NF = SD20.D2Doc,
                                  Emissao = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                                  Tipo = "REMESSA",
                                  Lote = SD20.D2Lotectl,
                                  Saldo = SD20.D2Quant - SD20.D2Qtdedev
                              }
                             ).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("PoderTerceiros");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "PoderTerceiros");

                sheet.Cells[1, 1].Value = "NF";
                sheet.Cells[1, 2].Value = "Serie";
                sheet.Cells[1, 3].Value = "Emissao";
                sheet.Cells[1, 4].Value = "Produto";
                sheet.Cells[1, 5].Value = "Lote";
                sheet.Cells[1, 6].Value = "Saldo";
                sheet.Cells[1, 7].Value = "Tipo";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.NF;
                    sheet.Cells[i, 2].Value = Pedido.Serie;
                    sheet.Cells[i, 3].Value = Pedido.Emissao;
                    sheet.Cells[i, 4].Value = Pedido.Produto;
                    sheet.Cells[i, 5].Value = Pedido.Lote;
                    sheet.Cells[i, 6].Value = Pedido.Saldo;
                    sheet.Cells[i, 7].Value = Pedido.Tipo;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PoderTerceiros.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "PoderTerceiros Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
