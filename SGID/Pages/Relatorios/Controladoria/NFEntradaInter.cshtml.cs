using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class NFEntradaInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioNFEntrada> Relatorio = new List<RelatorioNFEntrada>();

        public NFEntradaInterModel(TOTVSINTERContext context,ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        public string DataRel { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string Data)
        {
            try
            {
                DataRel = Data;

                var query = (from SD10 in Protheus.Sd1010s
                             join SF40 in Protheus.Sf4010s on SD10.D1Tes equals SF40.F4Codigo
                             where SD10.DELET != "*" && SF40.DELET != "*" && (SD10.D1Cf == "1908" || SD10.D1Cf == "1909" || SD10.D1Cf == "1949")
                             && SD10.D1Dtdigit == Data
                             select new
                             {
                                 Filial = SD10.D1Filial,
                                 Doc = SD10.D1Doc,
                                 Serie = SD10.D1Serie,
                                 Emissao = SD10.D1Emissao,
                                 DtDigit = SD10.D1Dtdigit,
                                 NFori = SD10.D1Nfori,
                                 Seriori = SD10.D1Seriori,
                                 Datori = SD10.D1Datori,
                                 Fornece = SD10.D1Fornece,
                                 Loja = SD10.D1Loja,
                                 Total = SD10.D1Total,
                                 ICMS = SD10.D1Valicm,
                                 Tipo = SF40.F4Texto.Contains("SIMB") ? "S" : "R",
                             }
                             ).ToList();

                if (query.Count != 0)
                {
                    query.ForEach(x =>
                    {
                        if (!Relatorio.Any(d => d.NFEntrada == x.Doc && x.NFori == d.NFOriginal))
                        {

                            var Iguais = query
                            .Where(c => c.Filial == x.Filial && c.Doc == x.Doc && c.Serie == x.Serie && c.Emissao == x.Emissao
                            && c.DtDigit == x.DtDigit && c.NFori == x.NFori && c.Seriori == x.Seriori && c.Datori == x.Datori
                            && c.Fornece == x.Fornece && c.Loja == x.Loja && c.Tipo == x.Tipo).ToList();

                            double total = 0;
                            double Icms = 0;
                            Iguais.ForEach(x =>
                            {
                                total += x.Total;
                                Icms += x.ICMS;
                            });

                            Relatorio.Add(new RelatorioNFEntrada
                            {
                                FILIAL = x.Filial,
                                NFOriginal = x.NFori,
                                SerieOriginal = x.Seriori,
                                EmissaoOrig = $"{x.Datori.Substring(6, 2)}/{x.Datori.Substring(4, 2)}/{x.Datori.Substring(0, 4)}",
                                EmissaoEnt = $"{x.Emissao.Substring(6, 2)}/{x.Emissao.Substring(4, 2)}/{x.Emissao.Substring(0, 4)}",
                                DigiEnt = $"{x.DtDigit.Substring(6, 2)}/{x.DtDigit.Substring(4, 2)}/{x.DtDigit.Substring(0, 4)}",
                                TipoEnt = x.Tipo,
                                NFEntrada = x.Doc,
                                SerieEntrada = x.Serie,
                                Fornecedor = x.Fornece,
                                Loja = x.Loja,
                                ICMS = Icms
                            });
                        }
                    });
                }
                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NFEntradaInter",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {

                var query = (from SD10 in Protheus.Sd1010s
                             join SF40 in Protheus.Sf4010s on SD10.D1Tes equals SF40.F4Codigo
                             where SD10.DELET != "*" && SF40.DELET != "*" && (SD10.D1Cf == "1908" || SD10.D1Cf == "1909" || SD10.D1Cf == "1949")
                             && SD10.D1Dtdigit == DataRel
                             select new
                             {
                                 Filial = SD10.D1Filial,
                                 Doc = SD10.D1Doc,
                                 Serie = SD10.D1Serie,
                                 Emissao = SD10.D1Emissao,
                                 DtDigit = SD10.D1Dtdigit,
                                 NFori = SD10.D1Nfori,
                                 Seriori = SD10.D1Seriori,
                                 Datori = SD10.D1Datori,
                                 Fornece = SD10.D1Fornece,
                                 Loja = SD10.D1Loja,
                                 Total = SD10.D1Total,
                                 ICMS = SD10.D1Valicm,
                                 Tipo = SF40.F4Texto.Contains("SIMB") ? "S" : "R",
                             }
                             ).ToList();

                if (query.Count != 0)
                {
                    query.ForEach(x =>
                    {
                        if (!Relatorio.Any(d => d.NFEntrada == x.Doc && x.NFori == d.NFOriginal))
                        {

                            var Iguais = query
                            .Where(c => c.Filial == x.Filial && c.Doc == x.Doc && c.Serie == x.Serie && c.Emissao == x.Emissao
                            && c.DtDigit == x.DtDigit && c.NFori == x.NFori && c.Seriori == x.Seriori && c.Datori == x.Datori
                            && c.Fornece == x.Fornece && c.Loja == x.Loja && c.Tipo == x.Tipo).ToList();

                            double total = 0;
                            double Icms = 0;
                            Iguais.ForEach(x =>
                            {
                                total += x.Total;
                                Icms += x.ICMS;
                            });

                            Relatorio.Add(new RelatorioNFEntrada
                            {
                                FILIAL = x.Filial,
                                NFOriginal = x.NFori,
                                SerieOriginal = x.Seriori,
                                EmissaoOrig = $"{x.Datori.Substring(6, 2)}/{x.Datori.Substring(4, 2)}/{x.Datori.Substring(0, 4)}",
                                EmissaoEnt = $"{x.Emissao.Substring(6, 2)}/{x.Emissao.Substring(4, 2)}/{x.Emissao.Substring(0, 4)}",
                                DigiEnt = $"{x.DtDigit.Substring(6, 2)}/{x.DtDigit.Substring(4, 2)}/{x.DtDigit.Substring(0, 4)}",
                                TipoEnt = x.Tipo,
                                NFEntrada = x.Doc,
                                SerieEntrada = x.Serie,
                                Fornecedor = x.Fornece,
                                Loja = x.Loja,
                                ICMS = Icms
                            });
                        }
                    });
                }

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("NFEntrada");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "NFEntrada");

                sheet.Cells[1, 1].Value = "FILIAL";
                sheet.Cells[1, 2].Value = "NF ORIGINAL";
                sheet.Cells[1, 3].Value = "SERIE ORIGINAL";
                sheet.Cells[1, 4].Value = "EMISSAO ORIG.";
                sheet.Cells[1, 5].Value = "EMISSAO ENT.";
                sheet.Cells[1, 6].Value = "DIGITAÇÃO ENT.";
                sheet.Cells[1, 7].Value = "TIPO ENT.";
                sheet.Cells[1, 8].Value = "NF ENTRADA";
                sheet.Cells[1, 9].Value = "SERIE ENTRADA";
                sheet.Cells[1, 10].Value = "FORNECEDOR";
                sheet.Cells[1, 11].Value = "LOJA";
                sheet.Cells[1, 12].Value = "ICMS";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.FILIAL;
                    sheet.Cells[i, 2].Value = Pedido.NFOriginal;
                    sheet.Cells[i, 3].Value = Pedido.SerieOriginal;
                    sheet.Cells[i, 4].Value = Pedido.EmissaoOrig;
                    sheet.Cells[i, 5].Value = Pedido.EmissaoEnt;
                    sheet.Cells[i, 6].Value = Pedido.DigiEnt;
                    sheet.Cells[i, 7].Value = Pedido.TipoEnt;
                    sheet.Cells[i, 8].Value = Pedido.NFEntrada;
                    sheet.Cells[i, 9].Value = Pedido.SerieEntrada;
                    sheet.Cells[i, 10].Value = Pedido.Fornecedor;
                    sheet.Cells[i, 11].Value = Pedido.Loja;
                    sheet.Cells[i, 12].Value = Pedido.ICMS;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "NFEntradaInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NFEntradaInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
