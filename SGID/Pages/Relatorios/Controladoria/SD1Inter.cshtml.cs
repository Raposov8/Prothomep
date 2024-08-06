using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Controladoria;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class SD1InterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public List<SD1> Relatorio { get; set; } = new List<SD1>();

        public SD1InterModel(TOTVSINTERContext denuo, ApplicationDbContext sgid)
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
                var query = (
                    from SD10 in Protheus.Sd1010s
                    join SF10 in Protheus.Sf1010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Doc, Serie = SD10.D1Serie, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF10.F1Filial, Doc = SF10.F1Doc, Serie = SF10.F1Serie, Fornece = SF10.F1Fornece, Loja = SF10.F1Loja }
                    join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                    join SA10 in Protheus.Sa1010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA10.A1Cod, Loja = SA10.A1Loja }
                    where SD10.DELET != "*" && SF10.DELET != "*" && SB10.DELET != "*" && SA10.DELET != "*"
                    && Tipo.Contains(SF10.F1Tipo) && (int)(object)SD10.D1Dtdigit >= (int)(object)DataI && (int)(object)SD10.D1Dtdigit <= (int)(object)DataF
                    select new SD1
                    {
                        Filial = SD10.D1Filial,
                        Produto = SD10.D1Cod,
                        DescProd = SB10.B1Desc,
                        QTDE = SD10.D1Quant,
                        VLUnit = SD10.D1Vunit,
                        Total = SD10.D1Total,
                        IPI = SD10.D1Valipi,
                        ICMS = SD10.D1Valicm,
                        TES = SD10.D1Tes,
                        CodFis = SD10.D1Cf,
                        Fornecedor = SD10.D1Fornece,
                        NomFor = SA10.A1Nome,
                        Documento = SD10.D1Doc,
                        Digitacao = SD10.D1Dtdigit,
                        Operacao = "ENTRADA",
                        NFOrig = SD10.D1Nfori,
                        SerieOrig = SD10.D1Seriori
                    });


                Relatorio = query.ToList();


                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SD1 Inter", user);

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

                var query = (from SD10 in Protheus.Sd1010s
                             join SF10 in Protheus.Sf1010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Doc, Serie = SD10.D1Serie, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF10.F1Filial, Doc = SF10.F1Doc, Serie = SF10.F1Serie, Fornece = SF10.F1Fornece, Loja = SF10.F1Loja }
                             join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                             join SA20 in Protheus.Sa2010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA20.A2Cod, Loja = SA20.A2Loja }
                             where SD10.DELET != "*" && SF10.DELET != "*" && SB10.DELET != "*" && SA20.DELET != "*"
                             && !Tipo.Contains(SF10.F1Tipo) && (int)(object)SD10.D1Dtdigit >= (int)(object)DataI && (int)(object)SD10.D1Dtdigit <= (int)(object)DataF
                             select new SD1
                             {
                                 Filial = SD10.D1Filial,
                                 Produto = SD10.D1Cod,
                                 DescProd = SB10.B1Desc,
                                 QTDE = SD10.D1Quant,
                                 VLUnit = SD10.D1Vunit,
                                 Total = SD10.D1Total,
                                 IPI = SD10.D1Valipi,
                                 ICMS = SD10.D1Valicm,
                                 TES = SD10.D1Tes,
                                 CodFis = SD10.D1Cf,
                                 Fornecedor = SD10.D1Fornece,
                                 NomFor = SA20.A2Nome,
                                 Documento = SD10.D1Doc,
                                 Digitacao = SD10.D1Dtdigit,
                                 Operacao = "ENTRADA",
                                 NFOrig = SD10.D1Nfori,
                                 SerieOrig = SD10.D1Seriori
                             }).Union(
                    from SD10 in Protheus.Sd1010s
                    join SF10 in Protheus.Sf1010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Doc, Serie = SD10.D1Serie, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF10.F1Filial, Doc = SF10.F1Doc, Serie = SF10.F1Serie, Fornece = SF10.F1Fornece, Loja = SF10.F1Loja }
                    join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                    join SA10 in Protheus.Sa1010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA10.A1Cod, Loja = SA10.A1Loja }
                    where SD10.DELET != "*" && SF10.DELET != "*" && SB10.DELET != "*" && SA10.DELET != "*"
                    && Tipo.Contains(SF10.F1Tipo) && (int)(object)SD10.D1Dtdigit >= (int)(object)DataI && (int)(object)SD10.D1Dtdigit <= (int)(object)DataF
                    select new SD1
                    {
                        Filial = SD10.D1Filial,
                        Produto = SD10.D1Cod,
                        DescProd = SB10.B1Desc,
                        QTDE = SD10.D1Quant,
                        VLUnit = SD10.D1Vunit,
                        Total = SD10.D1Total,
                        IPI = SD10.D1Valipi,
                        ICMS = SD10.D1Valicm,
                        TES = SD10.D1Tes,
                        CodFis = SD10.D1Cf,
                        Fornecedor = SD10.D1Fornece,
                        NomFor = SA10.A1Nome,
                        Documento = SD10.D1Doc,
                        Digitacao = SD10.D1Dtdigit,
                        Operacao = "ENTRADA",
                        NFOrig = SD10.D1Nfori,
                        SerieOrig = SD10.D1Seriori
                    });


                Relatorio = query.ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("SD1 Denuo");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "SD1 Denuo");

                sheet.Cells[1, 1].Value = "FILIAL";
                sheet.Cells[1, 2].Value = "COD. PROD.";
                sheet.Cells[1, 3].Value = "PRODUTO";
                sheet.Cells[1, 4].Value = "QTDE";
                sheet.Cells[1, 5].Value = "VL UNIT";
                sheet.Cells[1, 6].Value = "TOTAL";
                sheet.Cells[1, 7].Value = "IPI";
                sheet.Cells[1, 8].Value = "ICMS";
                sheet.Cells[1, 9].Value = "TES";
                sheet.Cells[1, 10].Value = "COD FIS";
                sheet.Cells[1, 11].Value = "COD. FOR";
                sheet.Cells[1, 12].Value = "FORNECEDOR";
                sheet.Cells[1, 13].Value = "DOCUMENTO";
                sheet.Cells[1, 14].Value = "DIGITAÇÃO";
                sheet.Cells[1, 15].Value = "OPERAÇÃO";
                sheet.Cells[1, 16].Value = "NF ORIG";
                sheet.Cells[1, 17].Value = "SERIE ORIG";

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
                    sheet.Cells[i, 9].Value = Pedido.TES;
                    sheet.Cells[i, 10].Value = Pedido.CodFis;
                    sheet.Cells[i, 11].Value = Pedido.Fornecedor;
                    sheet.Cells[i, 12].Value = Pedido.NomFor;
                    sheet.Cells[i, 13].Value = Pedido.Documento;
                    sheet.Cells[i, 14].Value = Pedido.Digitacao;
                    sheet.Cells[i, 15].Value = Pedido.Operacao;
                    sheet.Cells[i, 16].Value = Pedido.NFOrig;
                    sheet.Cells[i, 17].Value = Pedido.SerieOrig;

                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SD1Inter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SD1 Inter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
