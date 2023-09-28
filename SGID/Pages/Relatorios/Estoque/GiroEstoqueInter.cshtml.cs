using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Estoque;

namespace SGID.Pages.Relatorios.Estoque
{
    public class GiroEstoqueInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioGiroEstoque> Relatorio { get; set; } = new List<RelatorioGiroEstoque>();

        public GiroEstoqueInterModel(TOTVSINTERContext denuo, ApplicationDbContext sgid)
        {
            Protheus = denuo;
            SGID = sgid;
        }

        public void OnGet()
        {

            #region Dados
            var data = DateTime.Now;

            var Inicio = $"{data.Year - 1}{data.ToString("MM")}{data.ToString("dd")}";
            var Fim = data.ToString("yyyy/MM/dd").Replace("/", "");

            var Saida = (from SD20 in Protheus.Sd2010s
                         join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num }
                         join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                         where SC50.C5Utpoper == "F" && (int)(object)SD20.D2Emissao >= (int)(object)Inicio &&
                         (int)(object)SD20.D2Emissao <= (int)(object)Fim
                         select new
                         {
                             Codigo = SD20.D2Cod,
                             Desc = SB10.B1Desc,
                             Quant = SD20.D2Quant
                         }
                        ).GroupBy(x => new
                        {
                            x.Codigo,
                            x.Desc
                        }).Select(x => new
                        {
                            Codigo = x.Key.Codigo,
                            Desc = x.Key.Desc,
                            Quant = x.Sum(c => c.Quant)
                        })
                        .ToList();



            var Entrada = (from SB80 in Protheus.Sb8010s
                           join SB10 in Protheus.Sb1010s on SB80.B8Produto equals SB10.B1Cod
                           where SB80.DELET != "*" && (SB80.B8Local == "01" || SB80.B8Local == "02" || SB80.B8Local == "80")
                           select new
                           {
                               Codigo = SB80.B8Produto,
                               Desc = SB10.B1Desc,
                               Saldo = SB80.B8Saldo - SB80.B8Empenho,
                               Fabricante = SB10.B1Fabric,
                               Unidade = SB10.B1Xunnego == "000001" ? "OUTROS" : SB10.B1Xunnego == "000003" ? "COLUNA" : SB10.B1Xunnego == "000004" ? "NEURO" : SB10.B1Xunnego == "000005" ? "ORTOPEDIA" : SB10.B1Xunnego == "000006" ? "BUCOMAXILO" : SB10.B1Xunnego == "000007" ? "MATRIX" : SB10.B1Xunnego == "000008" ? "DENTAL" : SB10.B1Xunnego == "000009" ? "CMF" : ""
                           }
                           ).GroupBy(x => new
                           {
                               x.Codigo,
                               x.Desc,
                               x.Fabricante,
                               x.Unidade
                           }).Select(x => new RelatorioGiroEstoque
                           {
                               Codigo = x.Key.Codigo,
                               Desc = x.Key.Desc,
                               Saldo = x.Sum(c => c.Saldo),
                               Fabricante = x.Key.Fabricante,
                               Unidade = x.Key.Unidade
                           }).Where(x => x.Saldo >= 0).ToList();


            Entrada.ForEach(x =>
            {
                var produto = Saida.FirstOrDefault(c => c.Codigo.Trim() == x.Codigo.Trim());

                if (produto != null)
                {
                    x.Saida = produto.Quant;
                    x.Giro = produto.Quant / x.Saldo;
                    x.EstoqueMinimo = Math.Ceiling((produto.Quant / 12) * 3);
                    x.PontoPedido = x.Saldo - x.EstoqueMinimo <= 0 ? x.EstoqueMinimo : 0;
                }
            });

            Relatorio = Entrada;

            #endregion

        }

        public IActionResult OnPostExport()
        {
            try
            {
                #region Dados
                var data = DateTime.Now;

                var Inicio = $"{data.Year - 1}{data.ToString("MM")}{data.ToString("dd")}";
                var Fim = data.ToString("yyyy/MM/dd").Replace("/", "");

                var Saida = (from SD20 in Protheus.Sd2010s
                             join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             where SC50.C5Utpoper == "F" && (int)(object)SD20.D2Emissao >= (int)(object)Inicio &&
                             (int)(object)SD20.D2Emissao <= (int)(object)Fim
                             select new
                             {
                                 Codigo = SD20.D2Cod,
                                 Desc = SB10.B1Desc,
                                 Quant = SD20.D2Quant
                             }
                            ).GroupBy(x => new
                            {
                                x.Codigo,
                                x.Desc
                            }).Select(x => new
                            {
                                Codigo = x.Key.Codigo,
                                Desc = x.Key.Desc,
                                Quant = x.Sum(c => c.Quant)
                            })
                            .ToList();



                var Entrada = (from SB80 in Protheus.Sb8010s
                               join SB10 in Protheus.Sb1010s on SB80.B8Produto equals SB10.B1Cod
                               where SB80.DELET != "*" && (SB80.B8Local == "01" || SB80.B8Local == "02" || SB80.B8Local == "80")
                               select new
                               {
                                   Codigo = SB80.B8Produto,
                                   Desc = SB10.B1Desc,
                                   Saldo = SB80.B8Saldo - SB80.B8Empenho,
                                   Fabricante = SB10.B1Fabric,
                                   Unidade = SB10.B1Xunnego == "000001" ? "OUTROS" : SB10.B1Xunnego == "000003" ? "COLUNA" : SB10.B1Xunnego == "000004" ? "NEURO" : SB10.B1Xunnego == "000005" ? "ORTOPEDIA" : SB10.B1Xunnego == "000006" ? "BUCOMAXILO" : SB10.B1Xunnego == "000007" ? "MATRIX" : SB10.B1Xunnego == "000008" ? "DENTAL" : SB10.B1Xunnego == "000009" ? "CMF" : ""
                               }
                               ).GroupBy(x => new
                               {
                                   x.Codigo,
                                   x.Desc,
                                   x.Fabricante,
                                   x.Unidade
                               }).Select(x => new RelatorioGiroEstoque
                               {
                                   Codigo = x.Key.Codigo,
                                   Desc = x.Key.Desc,
                                   Saldo = x.Sum(c => c.Saldo),
                                   Fabricante = x.Key.Fabricante,
                                   Unidade = x.Key.Unidade
                               }).Where(x => x.Saldo >= 0).ToList();


                Entrada.ForEach(x =>
                {
                    var produto = Saida.FirstOrDefault(c => c.Codigo.Trim() == x.Codigo.Trim());

                    if (produto != null)
                    {
                        x.Saida = produto.Quant;
                        x.Giro = produto.Quant / x.Saldo;
                        x.EstoqueMinimo = Math.Ceiling((produto.Quant / 12) * 3);
                        x.PontoPedido = x.Saldo - x.EstoqueMinimo <= 0 ? x.EstoqueMinimo : 0;
                    }
                });

                Relatorio = Entrada;

                #endregion

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("GiroEstoque");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "GiroEstoque");

                sheet.Cells[1, 1].Value = "Codigo";
                sheet.Cells[1, 2].Value = "Fabricante";
                sheet.Cells[1, 3].Value = "Descrição";
                sheet.Cells[1, 4].Value = "Saldo Estoque";
                sheet.Cells[1, 5].Value = "Quant. Vendida ult 12 meses";
                sheet.Cells[1, 6].Value = "Giro";
                sheet.Cells[1, 7].Value = "Estoque minimo";
                sheet.Cells[1, 8].Value = "Ponto de Pedido";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Codigo;
                    sheet.Cells[i, 2].Value = Pedido.Fabricante;
                    sheet.Cells[i, 3].Value = Pedido.Desc;
                    sheet.Cells[i, 4].Value = Pedido.Saldo;
                    sheet.Cells[i, 5].Value = Pedido.Saida;
                    sheet.Cells[i, 6].Value = Pedido.Giro;
                    sheet.Cells[i, 7].Value = Pedido.EstoqueMinimo;
                    sheet.Cells[i, 8].Value = Pedido.PontoPedido;

                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "GiroEstoqueInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "GiroEstoqueInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
