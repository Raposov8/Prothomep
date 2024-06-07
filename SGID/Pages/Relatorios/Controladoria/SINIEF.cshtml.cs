using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Denuo;
using SGID.Models.DTO;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class SINIEFModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public int CurrentPage { get; set; }
        public int TotalPages { get; set; }
        public List<RelatorioSINIEF> Relatorio { get; set; } = new List<RelatorioSINIEF>();

        public SINIEFModel(TOTVSDENUOContext context,ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                var query = (from SD20 in Protheus.Sd2010s
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             where SD20.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC50.DELET != "*" && SF20.F2Tipo != "B"
                             && SF20.F2Tipo != "D" && (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "") &&
                             (SD20.D2Cf == "5908" || SD20.D2Cf == "5949" || SD20.D2Cf == "6908" || SD20.D2Cf == "6949")
                             select new SINIEF
                             {
                                 Filial = SD20.D2Filial,
                                 Doc = SD20.D2Doc,
                                 Serie = SD20.D2Serie,
                                 Emissao = SD20.D2Emissao,
                                 Clifor = SD20.D2Cliente,
                                 Loja = SD20.D2Loja,
                                 NomCliFor = SA10.A1Nome,
                                 Paciente = SC50.C5XNmpac,
                                 Valicm = SD20.D2Valicm,
                                 Operacao = "SAIDA",
                                 Pedido = SD20.D2Pedido,
                                 Agend = SC50.C5Unumage,
                                 ValSaida = SD20.D2Total,
                                 CFOP = SD20.D2Cf
                             }
                             ).GroupBy(x => new
                             {
                                 x.Filial,
                                 x.Doc,
                                 x.Serie,
                                 x.Emissao,
                                 x.Clifor,
                                 x.Loja,
                                 x.NomCliFor,
                                 x.Paciente,
                                 x.Pedido,
                                 x.Agend,
                                 x.CFOP
                             }).OrderBy(x => x.Key.Emissao);

                var queryselect = query.Select(x => new SINIEF
                {
                    Filial = x.Key.Filial,
                    Doc = x.Key.Doc,
                    Serie = x.Key.Serie,
                    Clifor = x.Key.Clifor,
                    Emissao = x.Key.Emissao,
                    Loja = x.Key.Loja,
                    NomCliFor = x.Key.NomCliFor,
                    Paciente = x.Key.Paciente,
                    Valicm = x.Sum(c => c.Valicm),
                    Pedido = x.Key.Pedido,
                    Agend = x.Key.Agend,
                    ValSaida = x.Sum(c => c.ValSaida),
                    CFOP = x.Key.CFOP
                });

                var Relatorios = queryselect.ToList();

                Relatorios.ForEach(obj =>
                {
                    var query = (from SD10 in Protheus.Sd1010s
                                 join SF40 in Protheus.Sf4010s on SD10.D1Tes equals SF40.F4Codigo
                                 where SD10.DELET != "*" && SF40.DELET != "*" &&
                                 obj.Filial == SD10.D1Filial && obj.Doc == SD10.D1Nfori && obj.Serie == SD10.D1Seriori && obj.Clifor == SD10.D1Fornece && obj.Loja == SD10.D1Loja
                                 select new SINIEFSUP
                                 {
                                     Filial = SD10.D1Filial,
                                     Doc = SD10.D1Doc,
                                     Serie = SD10.D1Serie,
                                     Emissao = SD10.D1Emissao,
                                     NFori = SD10.D1Nfori,
                                     Seriori = SD10.D1Seriori,
                                     Fornece = SD10.D1Fornece,
                                     Loja = SD10.D1Loja,
                                     Tipo = SF40.F4Texto.Contains("SIMB") ? "S" : "R",
                                     Valicm = SD10.D1Valicm,
                                     TotalDev = SD10.D1Total,
                                     CFOP = SD10.D1Cf
                                 }).GroupBy(x => new
                                 {
                                     x.Filial,
                                     x.Doc,
                                     x.Serie,
                                     x.Emissao,
                                     x.NFori,
                                     x.Seriori,
                                     x.Fornece,
                                     x.Loja,
                                     x.Tipo,
                                     x.CFOP
                                 }).Select(x => new SINIEFSUP
                                 {
                                     Filial = x.Key.Filial,
                                     Doc = x.Key.Doc,
                                     Serie = x.Key.Serie,
                                     Emissao = x.Key.Emissao,
                                     NFori = x.Key.NFori,
                                     Seriori = x.Key.Seriori,
                                     Fornece = x.Key.Fornece,
                                     Loja = x.Key.Loja,
                                     Tipo = x.Key.Tipo,
                                     Valicm = x.Sum(c => c.Valicm),
                                     TotalDev = x.Sum(c => c.TotalDev),
                                     CFOP = x.Key.CFOP
                                 }).ToList();

                    Relatorio.Add(new RelatorioSINIEF
                    {
                        Filial = obj.Filial,
                        Doc = obj.Doc,
                        Serie = obj.Serie,
                        Emissao = obj.Emissao,
                        CliFor = obj.Clifor,
                        Loja = obj.Loja,
                        NomCliFor = obj.NomCliFor,
                        Paciente = obj.Paciente,
                        Valicm = obj.Valicm,
                        Agend = obj.Agend,
                        ValSaida = obj.ValSaida,
                        Pedido = obj.Pedido,
                        CFOP = obj.CFOP
                    });

                    query.ForEach(banco =>
                    {
                        var linha = Relatorio.First(x => x.Filial == banco.Filial && x.Doc == banco.NFori && x.Serie == banco.Seriori && x.CliFor == banco.Fornece && x.Loja == banco.Loja);

                        if (banco.Tipo == "S")
                        {
                            linha.NFSimb = banco.Doc;
                            linha.Emissimb = banco.Emissao;
                            linha.ValRetReal += banco.TotalDev;
                        }
                        else
                        {
                            linha.NFReal = banco.Doc;
                            linha.Emisreal = banco.Emissao;
                            linha.ValRetReal += banco.TotalDev;
                        }

                        if (linha.NFSimb == null && linha.NFReal == null) linha.Dias = (Convert.ToDateTime($"{linha.Emissao.Substring(6, 2)}/{linha.Emissao.Substring(4, 2)}/{linha.Emissao.Substring(0, 4)}") - DateTime.Now).TotalDays;

                        linha.ValicmRet = banco.Valicm;
                    });
                });

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SINIEF",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {

                #region Relatorio
                var query = (from SD20 in Protheus.Sd2010s
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             where SD20.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC50.DELET != "*" && SF20.F2Tipo != "B"
                             && SF20.F2Tipo != "D" && (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "") &&
                             (SD20.D2Cf == "5908" || SD20.D2Cf == "5949" || SD20.D2Cf == "6908" || SD20.D2Cf == "6949")
                             select new SINIEF
                             {
                                 Filial = SD20.D2Filial,
                                 Doc = SD20.D2Doc,
                                 Serie = SD20.D2Serie,
                                 Emissao = SD20.D2Emissao,
                                 Clifor = SD20.D2Cliente,
                                 Loja = SD20.D2Loja,
                                 NomCliFor = SA10.A1Nome,
                                 Paciente = SC50.C5XNmpac,
                                 Valicm = SD20.D2Valicm,
                                 Operacao = "SAIDA",
                                 Pedido = SD20.D2Pedido,
                                 Agend = SC50.C5Unumage,
                                 ValSaida = SD20.D2Total,
                                 CFOP = SD20.D2Cf
                             }
                              ).GroupBy(x => new
                              {
                                  x.Filial,
                                  x.Doc,
                                  x.Serie,
                                  x.Emissao,
                                  x.Clifor,
                                  x.Loja,
                                  x.NomCliFor,
                                  x.Paciente,
                                  x.Pedido,
                                  x.Agend,
                                  x.CFOP
                              }).OrderBy(x => x.Key.Emissao);

                var queryselect = query.Select(x => new SINIEF
                {
                    Filial = x.Key.Filial,
                    Doc = x.Key.Doc,
                    Serie = x.Key.Serie,
                    Clifor = x.Key.Clifor,
                    Emissao = x.Key.Emissao,
                    Loja = x.Key.Loja,
                    NomCliFor = x.Key.NomCliFor,
                    Paciente = x.Key.Paciente,
                    Valicm = x.Sum(c => c.Valicm),
                    Pedido = x.Key.Pedido,
                    Agend = x.Key.Agend,
                    ValSaida = x.Sum(c => c.ValSaida),
                    CFOP = x.Key.CFOP
                });

                var Relatorios = queryselect.ToList();

                Relatorios.ForEach(obj =>
                {
                    var query = (from SD10 in Protheus.Sd1010s
                                 join SF40 in Protheus.Sf4010s on SD10.D1Tes equals SF40.F4Codigo
                                 where SD10.DELET != "*" && SF40.DELET != "*" &&
                                 obj.Filial == SD10.D1Filial && obj.Doc == SD10.D1Nfori && obj.Serie == SD10.D1Seriori && obj.Clifor == SD10.D1Fornece && obj.Loja == SD10.D1Loja
                                 select new SINIEFSUP
                                 {
                                     Filial = SD10.D1Filial,
                                     Doc = SD10.D1Doc,
                                     Serie = SD10.D1Serie,
                                     Emissao = SD10.D1Emissao,
                                     NFori = SD10.D1Nfori,
                                     Seriori = SD10.D1Seriori,
                                     Fornece = SD10.D1Fornece,
                                     Loja = SD10.D1Loja,
                                     Tipo = SF40.F4Texto.Contains("SIMB") ? "S" : "R",
                                     Valicm = SD10.D1Valicm,
                                     TotalDev = SD10.D1Total,
                                     CFOP = SD10.D1Cf
                                 }).GroupBy(x => new
                                 {
                                     x.Filial,
                                     x.Doc,
                                     x.Serie,
                                     x.Emissao,
                                     x.NFori,
                                     x.Seriori,
                                     x.Fornece,
                                     x.Loja,
                                     x.Tipo,
                                     x.CFOP
                                 }).Select(x => new SINIEFSUP
                                 {
                                     Filial = x.Key.Filial,
                                     Doc = x.Key.Doc,
                                     Serie = x.Key.Serie,
                                     Emissao = x.Key.Emissao,
                                     NFori = x.Key.NFori,
                                     Seriori = x.Key.Seriori,
                                     Fornece = x.Key.Fornece,
                                     Loja = x.Key.Loja,
                                     Tipo = x.Key.Tipo,
                                     Valicm = x.Sum(c => c.Valicm),
                                     TotalDev = x.Sum(c => c.TotalDev),
                                     CFOP = x.Key.CFOP
                                 }).ToList();

                    Relatorio.Add(new RelatorioSINIEF
                    {
                        Filial = obj.Filial,
                        Doc = obj.Doc,
                        Serie = obj.Serie,
                        Emissao = obj.Emissao,
                        CliFor = obj.Clifor,
                        Loja = obj.Loja,
                        NomCliFor = obj.NomCliFor,
                        Paciente = obj.Paciente,
                        Valicm = obj.Valicm,
                        Agend = obj.Agend,
                        ValSaida = obj.ValSaida,
                        Pedido = obj.Pedido,
                        CFOP = obj.CFOP
                    });

                    query.ForEach(banco =>
                    {
                        var linha = Relatorio.First(x => x.Filial == banco.Filial && x.Doc == banco.NFori && x.Serie == banco.Seriori && x.CliFor == banco.Fornece && x.Loja == banco.Loja);

                        if (banco.Tipo == "S")
                        {
                            linha.NFSimb = banco.Doc;
                            linha.Emissimb = banco.Emissao;
                            linha.ValRetReal += banco.TotalDev;
                        }
                        else
                        {
                            linha.NFReal = banco.Doc;
                            linha.Emisreal = banco.Emissao;
                            linha.ValRetReal += banco.TotalDev;
                        }

                        if (linha.NFSimb == null && linha.NFReal == null) linha.Dias = (Convert.ToDateTime($"{linha.Emissao.Substring(6, 2)}/{linha.Emissao.Substring(4, 2)}/{linha.Emissao.Substring(0, 4)}") - DateTime.Now).TotalDays;

                        linha.ValicmRet = banco.Valicm;
                    });
                });
                #endregion


                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("SINIEF");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "SINIEF");

                sheet.Cells[1, 1].Value = "FILIAL";
                sheet.Cells[1, 2].Value = "DOC";
                sheet.Cells[1, 3].Value = "SERIE";
                sheet.Cells[1, 4].Value = "EMISSAO";
                sheet.Cells[1, 5].Value = "CLIFOR";
                sheet.Cells[1, 6].Value = "LOJA";
                sheet.Cells[1, 7].Value = "NOMCLIFOR";
                sheet.Cells[1, 8].Value = "PACIENTE";
                sheet.Cells[1, 9].Value = "VALICM";
                sheet.Cells[1, 10].Value = "NFSIMB";
                sheet.Cells[1, 11].Value = "NFREAL";
                sheet.Cells[1, 12].Value = "EMISSIMB";
                sheet.Cells[1, 13].Value = "EMISREAL";
                sheet.Cells[1, 14].Value = "VALICMRET";
                sheet.Cells[1, 15].Value = "DIAS";
                sheet.Cells[1, 16].Value = "PEDIDO";
                sheet.Cells[1, 17].Value = "AGEND";
                sheet.Cells[1, 18].Value = "VALSAIDA";
                sheet.Cells[1, 19].Value = "VALRETSIMB";
                sheet.Cells[1, 20].Value = "VALRETREAL";
                sheet.Cells[1, 21].Value = "CFOP";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.Doc;
                    sheet.Cells[i, 3].Value = Pedido.Serie;
                    sheet.Cells[i, 4].Value = Pedido.Emissao;
                    sheet.Cells[i, 5].Value = Pedido.CliFor;
                    sheet.Cells[i, 6].Value = Pedido.Loja;
                    sheet.Cells[i, 7].Value = Pedido.NomCliFor;
                    sheet.Cells[i, 8].Value = Pedido.Paciente;
                    
                    sheet.Cells[i, 9].Value = Pedido.Valicm;
                    sheet.Cells[i, 9].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 10].Value = Pedido.NFSimb;
                    sheet.Cells[i, 11].Value = Pedido.NFReal;
                    sheet.Cells[i, 12].Value = Pedido.Emissimb;
                    sheet.Cells[i, 13].Value = Pedido.Emisreal;

                    sheet.Cells[i, 14].Value = Pedido.ValicmRet;
                    sheet.Cells[i, 14].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 15].Value = Pedido.Dias;
                    sheet.Cells[i, 16].Value = Pedido.Pedido;
                    sheet.Cells[i, 17].Value = Pedido.Agend;

                    sheet.Cells[i, 18].Value = Pedido.ValSaida;
                    sheet.Cells[i, 18].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 19].Value = Pedido.ValRetSimb;
                    sheet.Cells[i, 19].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 20].Value = Pedido.ValRetReal;
                    sheet.Cells[i, 20].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 21].Value = Pedido.CFOP;

                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SINIEF.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SINIEF Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
