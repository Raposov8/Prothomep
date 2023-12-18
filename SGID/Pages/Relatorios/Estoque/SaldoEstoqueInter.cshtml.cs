using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Estoque;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.Estoque
{
    public class SaldoEstoqueInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioSaldoEstoque> Relatorio { get; set; } = new List<RelatorioSaldoEstoque>();
        public List<string> Locals { get; set; } = new List<string>();
        public List<string> Locais { get; set; } = new List<string> { "01", "80", "70", "30" };

        public List<string> Produtos { get; set; } = new List<string>();

        public SaldoEstoqueInterModel(TOTVSINTERContext protheus, ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }

        public void OnGet()
        {
            try
            {

                Relatorio = (from SB20 in Protheus.Sb2010s
                             join SB10 in Protheus.Sb1010s on SB20.B2Cod equals SB10.B1Cod
                             where SB10.DELET != "*" && SB20.DELET != "*" && SB20.B2Qatu != 0
                             && SB20.B2Filial == "01" && Locais.Contains(SB20.B2Local)
                             select new RelatorioSaldoEstoque
                             {
                                 Filial = SB20.B2Filial,
                                 Produto = SB20.B2Cod,
                                 DescProd = SB10.B1Desc,
                                 Local = SB20.B2Local,
                                 Saldo = SB20.B2Qatu,
                                 UM = SB10.B1Um,
                                 Empenho = SB20.B2Reserva,
                             }).ToList();

                Produtos = Protheus.Sb2010s.Where(x => x.DELET != "*" && x.B2Qatu != 0
                              && x.B2Filial == "01").Select(x => x.B2Cod).Distinct().ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SaldoEstoque Inter", user);
            }
        }

        public IActionResult OnPost(string Local, string Produto)
        {
            try
            {

                Relatorio = (from SB20 in Protheus.Sb2010s
                             join SB10 in Protheus.Sb1010s on SB20.B2Cod equals SB10.B1Cod
                             where SB10.DELET != "*" && SB20.DELET != "*" && SB20.B2Qatu != 0
                             && SB20.B2Filial == "01" && Locais.Contains(SB20.B2Local)
                             select new RelatorioSaldoEstoque
                             {
                                 Filial = SB20.B2Filial,
                                 Produto = SB20.B2Cod,
                                 DescProd = SB10.B1Desc,
                                 Local = SB20.B2Local,
                                 Saldo = SB20.B2Qatu,
                                 UM = SB10.B1Um,
                                 Empenho = SB20.B2Reserva,
                             }).ToList();

                if (Local != "" && Local != null)
                {
                    Relatorio = Relatorio.Where(x => x.Local == Local).ToList();
                }

                if (Produto != "" && Produto != null)
                {
                    Relatorio = Relatorio.Where(x => x.Produto.Trim() == Produto.Trim()).ToList();
                }

                Produtos = Protheus.Sb2010s.Where(x => x.DELET != "*" && x.B2Qatu != 0
                             && x.B2Filial == "01").Select(x => x.B2Cod).Distinct().ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SaldoEstoque Inter", user);
                return Page();
            }
        }

        public IActionResult OnPostExport(string Local, string Produto)
        {
            try
            {
                Relatorio = (from SB20 in Protheus.Sb2010s
                             join SB10 in Protheus.Sb1010s on SB20.B2Cod equals SB10.B1Cod
                             where SB10.DELET != "*" && SB20.DELET != "*" && SB20.B2Qatu != 0
                             && SB20.B2Filial == "01" && Locais.Contains(SB20.B2Local)
                             select new RelatorioSaldoEstoque
                             {
                                 Filial = SB20.B2Filial,
                                 Produto = SB20.B2Cod,
                                 DescProd = SB10.B1Desc,
                                 Local = SB20.B2Local,
                                 Saldo = SB20.B2Qatu,
                                 UM = SB10.B1Um,
                                 Empenho = SB20.B2Reserva,
                             }).ToList();

                if (Local != "" && Local != null)
                {
                    Relatorio = Relatorio.Where(x => x.Local == Local).ToList();
                }

                if (Produto != "" && Produto != null)
                {
                    Relatorio = Relatorio.Where(x => x.Produto.Trim() == Produto.Trim()).ToList();
                }

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("SaldoEstoque");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "SaldoEstoque");

                sheet.Cells[1, 1].Value = "Filial";
                sheet.Cells[1, 2].Value = "Produto";
                sheet.Cells[1, 3].Value = "Desc Prod.";
                sheet.Cells[1, 4].Value = "Local";
                sheet.Cells[1, 5].Value = "Saldo";
                sheet.Cells[1, 6].Value = "Empenho";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.Produto;
                    sheet.Cells[i, 3].Value = Pedido.DescProd;
                    sheet.Cells[i, 4].Value = Pedido.Local;
                    sheet.Cells[i, 5].Value = Pedido.Saldo;
                    sheet.Cells[i, 6].Value = Pedido.Empenho;

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
