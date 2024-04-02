using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;

namespace SGID.Pages.Relatorios
{
    [Authorize]
    public class SubDistribuidorNFaturadoInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioSubDistribuidor> Relatorio { get; set; } = new List<RelatorioSubDistribuidor>();
        public double Total { get; set; }
        public double Desc { get; set; }
        public string Nreduz { get; set; }

        public SubDistribuidorNFaturadoInterModel(TOTVSINTERContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }

        public void OnGet()
        {
            try
            {
                var lista = (from SC50 in Protheus.Sc5010s
                             join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                             join SB10 in Protheus.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SC50.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC60.DELET != "*" &&
                             SA30.DELET != "*" && SB10.DELET != "*" && SA10.A1Clinter == "S" && SC50.C5Nota == "" &&
                             SC60.C6Qtdven - SC60.C6Qtdent != 0
                             && SA10.A1Cgc.Substring(0, 8) != "04715053"
                             orderby SA10.A1Nome, SC50.C5Emissao
                             select new RelatorioSubDistribuidor
                             {
                                 Num = SC50.C5Num,
                                 Emissao = $"{SC50.C5Emissao.Substring(6, 2)}/{SC50.C5Emissao.Substring(4, 2)}/{SC50.C5Emissao.Substring(0, 4)}",
                                 Nome = SA30.A3Nome,
                                 Desc = SB10.B1Desc,
                                 Quant = SC60.C6Qtdven - SC60.C6Qtdent,
                                 Doc = SC60.C6Produto,
                                 Total = (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven,
                                 Descon = SC60.C6Valdesc,
                                 Nreduz = SA10.A1Nome,
                                 Utpoper = SC50.C5Liberok,
                                 Fabricante = SB10.B1Fabric
                             }
                            ).ToList();

                Relatorio = lista;

                lista.ForEach(x => Total += x.Total);
                lista.ForEach(x => Desc += x.Descon);
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SubDistribuidorNFaturadoInter",user);
            }
        }

        public IActionResult OnPostAsync(string NReduz)
        {
            try
            {
                Nreduz = NReduz;

                var lista = (from SC50 in Protheus.Sc5010s
                             join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                             join SB10 in Protheus.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SC50.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC60.DELET != "*" &&
                             SA30.DELET != "*" && SB10.DELET != "*" && SA10.A1Clinter == "S" && SC50.C5Nota == "" &&
                             SC60.C6Qtdven - SC60.C6Qtdent != 0
                             && SA10.A1Cgc.Substring(0, 8) != "04715053"
                             orderby SA10.A1Nome, SC50.C5Emissao, SC50.C5Num, SB10.B1Desc
                             select new RelatorioSubDistribuidor
                             {
                                 Num = SC50.C5Num,
                                 Emissao = $"{SC50.C5Emissao.Substring(6, 2)}/{SC50.C5Emissao.Substring(4, 2)}/{SC50.C5Emissao.Substring(0, 4)}",
                                 Nome = SA30.A3Nome,
                                 Desc = SB10.B1Desc,
                                 Quant = SC60.C6Qtdven - SC60.C6Qtdent,
                                 Doc = SC60.C6Produto,
                                 Total = (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven,
                                 Descon = SC60.C6Valdesc,
                                 Nreduz = SA10.A1Nome,
                                 Utpoper = SC50.C5Liberok,
                                 Fabricante = SB10.B1Fabric
                             }
                            ).ToList();

                if (!string.IsNullOrEmpty(NReduz) && !string.IsNullOrWhiteSpace(NReduz))
                {
                    lista = lista.Where(x => x.Nreduz == NReduz).ToList();
                }

                Relatorio = lista;

                lista.ForEach(x => Total += x.Total);
                lista.ForEach(x => Desc += x.Descon);

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SubDistribuidorNFaturadoInter Post",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(string NReduz)
        {
            try
            {
                var lista = (from SC50 in Protheus.Sc5010s
                             join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                             join SB10 in Protheus.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SC50.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC60.DELET != "*" &&
                             SA30.DELET != "*" && SA10.A1Clinter == "S" && SC50.C5Nota == "" &&
                             SC60.C6Qtdven - SC60.C6Qtdent != 0
                             && SA10.A1Cgc.Substring(0, 8) != "04715053"
                             orderby SA10.A1Nome, SC50.C5Emissao, SC50.C5Num, SB10.B1Desc
                             select new RelatorioSubDistribuidor
                             {
                                 Num = SC50.C5Num,
                                 Emissao = SC50.C5Emissao,
                                 Nome = SA30.A3Nome,
                                 Desc = SB10.B1Desc,
                                 Quant = SC60.C6Qtdven - SC60.C6Qtdent,
                                 Doc = SC60.C6Produto,
                                 Total = (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven,
                                 Descon = SC60.C6Valdesc,
                                 Nreduz = SA10.A1Nome,
                                 Utpoper = SC50.C5Liberok,
                                 Fabricante = SB10.B1Fabric
                             }
                            ).ToList();

                if (!string.IsNullOrEmpty(NReduz) && !string.IsNullOrWhiteSpace(NReduz))
                {
                    lista = lista.Where(x => x.Nreduz == NReduz).ToList();
                }

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("SubDistribuidor Não Faturado");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "SubDistribuidor Não Faturado");

                sheet.Cells[1, 1].Value = "C5_NUM";
                sheet.Cells[1, 2].Value = "C5_EMISSAO";
                sheet.Cells[1, 3].Value = "A1_NOME";
                sheet.Cells[1, 4].Value = "C6_PRODUTO";
                sheet.Cells[1, 5].Value = "QTD";
                sheet.Cells[1, 6].Value = "B1_DESC";
                sheet.Cells[1, 7].Value = "Valor";
                sheet.Cells[1, 8].Value = "C6_VALDESC";
                sheet.Cells[1, 9].Value = "A3_NOME";
                sheet.Cells[1, 10].Value = "C5_LIBEROK";
                sheet.Cells[1, 11].Value = "Fabricante";



                int i = 2;

                lista.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Num;
                    sheet.Cells[i, 2].Value = $"{Pedido.Emissao.Substring(6, 2)}/{Pedido.Emissao.Substring(4, 2)}/{Pedido.Emissao.Substring(0, 4)}";
                    sheet.Cells[i, 3].Value = Pedido.Nreduz;
                    sheet.Cells[i, 4].Value = $"{Pedido.Doc}";
                    sheet.Cells[i, 5].Value = Pedido.Quant;
                    sheet.Cells[i, 6].Value = Pedido.Desc;
                    sheet.Cells[i, 7].Value = Pedido.Total;
                    sheet.Cells[i, 8].Value = Pedido.Descon;
                    sheet.Cells[i, 9].Value = Pedido.Nome;
                    sheet.Cells[i, 10].Value = Pedido.Utpoper;
                    sheet.Cells[i, 11].Value = Pedido.Fabricante;

                    i++;
                });

                var format = sheet.Cells[i, 4];
                format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SubDistribuidorNaoFaturadoInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SubDistribuidorNFaturadoInter Excel",user);

                return LocalRedirect("/error");
            }
        }
    }
}
