using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Controladoria;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class SD2InterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }


        public List<SD2> Relatorio { get; set; } = new List<SD2>();

        public SD2InterModel(TOTVSINTERContext denuo, ApplicationDbContext sgid)
        {
            Protheus = denuo;
            SGID = sgid;
        }
        public void OnGet()
        {
        }

        public IActionResult OnPost(DateTime DataInicio, DateTime DataFim)
        {
            Inicio = DataInicio;
            Fim = DataFim;
            string[] Tipo = new[] { "B", "D" };
            var DataI = DataInicio.ToString("yyyy/MM/dd").Replace("/", "");
            var DataF = DataFim.ToString("yyyy/MM/dd").Replace("/", "");

            try
            {
                var query = (from SD20 in Protheus.Sd2010s
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SA10 in Protheus.Sa1010s on new { Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             where SD20.DELET != "*" && SF20.DELET != "*" && SB10.DELET != "*" && SA10.DELET != "*"
                             && !Tipo.Contains(SF20.F2Tipo) && (int)(object)SD20.D2Emissao >= (int)(object)DataI && (int)(object)SD20.D2Emissao <= (int)(object)DataF
                             select new SD2
                             {
                                 Filial = SD20.D2Filial,
                                 Produto = SD20.D2Cod,
                                 DescProd = SB10.B1Desc,
                                 QTDE = SD20.D2Quant,
                                 VLUnit = SD20.D2Prcven,
                                 Total = SD20.D2Total,
                                 IPI = SD20.D2Valipi,
                                 ICMS = SD20.D2Valicm,
                                 TES = SD20.D2Tes,
                                 CodFis = SD20.D2Cf,
                                 Cliente = SD20.D2Cliente,
                                 NomCli = SA10.A1Nome,
                                 Documento = SD20.D2Doc,
                                 Serie = SD20.D2Serie,
                                 Custo = SD20.D2Custo1,
                                 VALIMP5 = SD20.D2Valimp5,
                                 VALIMP6 = SD20.D2Valimp6,
                                 ICMSDIF = SD20.D2Icmsdif,
                                 VALFECP = SD20.D2Valfecp,
                                 OPERACAO = "SAIDA"
                             })
                             ;


                Relatorio = query.ToList();


                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SD2 Inter", user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                string[] Tipo = new[] { "B", "D" };
                var DataI = DataInicio.ToString("yyyy/MM/dd").Replace("/", "");
                var DataF = DataFim.ToString("yyyy/MM/dd").Replace("/", "");

                var query = (from SD20 in Protheus.Sd2010s
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SA10 in Protheus.Sa1010s on new { Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             where SD20.DELET != "*" && SF20.DELET != "*" && SB10.DELET != "*" && SA10.DELET != "*"
                             && !Tipo.Contains(SF20.F2Tipo) && (int)(object)SD20.D2Emissao >= (int)(object)DataI && (int)(object)SD20.D2Emissao <= (int)(object)DataF
                             select new SD2
                             {
                                 Filial = SD20.D2Filial,
                                 Produto = SD20.D2Cod,
                                 DescProd = SB10.B1Desc,
                                 QTDE = SD20.D2Quant,
                                 VLUnit = SD20.D2Prcven,
                                 Total = SD20.D2Total,
                                 IPI = SD20.D2Valipi,
                                 ICMS = SD20.D2Valicm,
                                 TES = SD20.D2Tes,
                                 CodFis = SD20.D2Cf,
                                 Cliente = SD20.D2Cliente,
                                 NomCli = SA10.A1Nome,
                                 Documento = SD20.D2Doc,
                                 Serie = SD20.D2Serie,
                                 Custo = SD20.D2Custo1,
                                 VALIMP5 = SD20.D2Valimp5,
                                 VALIMP6 = SD20.D2Valimp6,
                                 ICMSDIF = SD20.D2Icmsdif,
                                 VALFECP = SD20.D2Valfecp,
                                 OPERACAO = "SAIDA"
                             });


                Relatorio = query.ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("SD2 Denuo");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "SD2 Denuo");

                sheet.Cells[1, 1].Value = "FILIAL";
                sheet.Cells[1, 2].Value = "COD. PROD.";
                sheet.Cells[1, 3].Value = "PRODUTO";
                sheet.Cells[1, 4].Value = "QTDE";
                sheet.Cells[1, 5].Value = "VL UNIT";
                sheet.Cells[1, 6].Value = "TOTAL";
                sheet.Cells[1, 7].Value = "IPI";
                sheet.Cells[1, 8].Value = "ICMS";
                sheet.Cells[1, 9].Value = "CUSTO";
                sheet.Cells[1, 10].Value = "VALIMP5";
                sheet.Cells[1, 11].Value = "VALIMP6";
                sheet.Cells[1, 12].Value = "ICMSDIF";
                sheet.Cells[1, 13].Value = "VALFECP";
                sheet.Cells[1, 14].Value = "TES";
                sheet.Cells[1, 15].Value = "COD FIS";
                sheet.Cells[1, 16].Value = "COD. CLI";
                sheet.Cells[1, 17].Value = "CLIENTE";
                sheet.Cells[1, 18].Value = "DOCUMENTO";
                sheet.Cells[1, 19].Value = "SERIE";
                sheet.Cells[1, 20].Value = "OPERAÇÃO";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.Produto;
                    sheet.Cells[i, 3].Value = Pedido.DescProd;
                    sheet.Cells[i, 4].Value = Pedido.QTDE;
                    sheet.Cells[i, 5].Value = Pedido.VLUnit;
                    sheet.Cells[i, 5].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 6].Value = Pedido.Total;
                    sheet.Cells[i, 6].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 7].Value = Pedido.IPI;
                    sheet.Cells[i, 7].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 8].Value = Pedido.ICMS;
                    sheet.Cells[i, 8].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 9].Value = Pedido.Custo;
                    sheet.Cells[i, 9].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 10].Value = Pedido.VALIMP5;
                    sheet.Cells[i, 10].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 11].Value = Pedido.VALIMP6;
                    sheet.Cells[i, 11].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 12].Value = Pedido.ICMSDIF;
                    sheet.Cells[i, 12].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 13].Value = Pedido.VALFECP;
                    sheet.Cells[i, 13].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 14].Value = Pedido.TES;
                    sheet.Cells[i, 15].Value = Pedido.CodFis;
                    sheet.Cells[i, 16].Value = Pedido.Cliente;
                    sheet.Cells[i, 17].Value = Pedido.NomCli;
                    sheet.Cells[i, 18].Value = Pedido.Documento;
                    sheet.Cells[i, 19].Value = Pedido.Serie;
                    sheet.Cells[i, 20].Value = Pedido.OPERACAO;

                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SD2Inter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SD2 Inter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
