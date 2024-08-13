using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Account.RH;
using SGID.Models.Denuo;
using SGID.Models.Diretoria;
using SGID.Models.Financeiro;
using SGID.Models.RH;

namespace SGID.Pages.DashBoards
{
    [Authorize]
    public class DashBoardDiretorio2Model : PageModel
    {
        public List<ValoresEmAberto> ValoresEmAberto { get; set; } = new List<ValoresEmAberto>();
        public List<ValoresEmAberto> Faturamento { get; set; } = new List<ValoresEmAberto>();

        public List<ValoresEmAberto> Baixa { get; set; } = new List<ValoresEmAberto>();

        public TOTVSDENUOContext ProtheusDenuo { get; set; }
        public TOTVSINTERContext ProtheusInter { get; set; }

        public string Mes { get; set; }
        public string Ano { get; set; }
        public ApplicationDbContext SGID { get; set; }

        public DashBoardDiretorio2Model(TOTVSDENUOContext denuo, TOTVSINTERContext inter, ApplicationDbContext sgid)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
            SGID = sgid;
        }

        public void OnGet()
        {
            
        }


        public JsonResult OnPostFaturado(string Tipo)
        {
            try
            {
                #region Parametros
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                DateTime data;
                string mesInicio = "";
                string anoInicio = "";
                string DataInicio = $"";

                DateTime Fim;

                string mesFim = "";
                string anoFim = "";
                string DataFim = $"";

                string[] CF = new string[] { "5551", "6551", "6107", "6109" };

                switch (Tipo)
                {
                    case "30":
                        {
                            data = DateTime.Now;
                            mesInicio = data.Month.ToString("D2");
                            anoInicio = data.Year.ToString();
                            DataInicio = $"{anoInicio}{mesInicio}01";

                            DataFim = $"{anoInicio}{mesInicio}31";

                            break;

                        }
                    case "60":
                        {
                            data = DateTime.Now.AddMonths(11);
                            mesInicio = data.Month.ToString("D2");
                            anoInicio = (data.Year - 1).ToString();
                            DataInicio = $"{anoInicio}{mesInicio}01";

                            Fim = DateTime.Now;

                            mesFim = Fim.Month.ToString("D2");
                            anoFim = Fim.Year.ToString();
                            DataFim = $"{anoFim}{mesFim}31";

                            break;
                        }
                    default:
                        {
                            
                            data = DateTime.Now.AddMonths(10);
                            mesInicio = data.Month.ToString("D2");
                            anoInicio = (data.Year - 1).ToString();
                            DataInicio = $"{anoInicio}{mesInicio}01";

                            Fim = DateTime.Now;

                            mesFim = Fim.Month.ToString("D2");
                            anoFim = Fim.Year.ToString();
                            DataFim = $"{anoFim}{mesFim}31";

                            break;
                        }
                }
                #endregion


                #region Faturados

                #region Intermedic

                #region Faturado
                var query = (from SD20 in ProtheusInter.Sd2010s
                             join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                             && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                             (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains(SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                             && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                             select new
                             {
                                 Login = SA30.A3Xlogin,
                                 NF = SD20.D2Doc,
                                 Total = SD20.D2Total,

                             });

                var resultadoInter = query.GroupBy(x => new
                {
                    x.Login,
                    x.NF,
                    x.Total
                }).Select(x => new RelatorioCirurgiasFaturadas
                {
                    A3Nome = x.Key.Login.Trim(),
                    Nf = x.Key.NF,
                    Total = x.Sum(c => c.Total),
                }).ToList();
                #endregion

                #region Devolucao
                var DevolInter = (from SD10 in ProtheusInter.Sd1010s
                                  join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                  join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                  join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                  where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                  && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                  && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                  && (int)(object)SD10.D1Dtdigit >= 20200801 && SC50.C5Utpoper == "F"
                                  orderby SA30.A3Nome
                                  select new
                                  {
                                      SD10.D1Filial,
                                      SD10.D1Fornece,
                                      SD10.D1Loja,
                                      SA10.A1Nome,
                                      SA10.A1Clinter,
                                      SD10.D1Doc,
                                      SD10.D1Serie,
                                      SD10.D1Dtdigit,
                                      SD10.D1Total,
                                      SD10.D1Valdesc,
                                      SD10.D1Valipi,
                                      SD10.D1Valicm,
                                      SA30.A3Nome,
                                      SD10.D1Nfori,
                                      SD10.D1Seriori,
                                      SD10.D1Datori,
                                      SD10.D1Emissao
                                  }
                             )
                             .GroupBy(x => new
                             {
                                 x.A1Nome,
                                 x.A3Nome,
                                 x.D1Filial,
                                 x.D1Fornece,
                                 x.D1Loja,
                                 x.A1Clinter,
                                 x.D1Doc,
                                 x.D1Serie,
                                 x.D1Emissao,
                                 x.D1Dtdigit,
                                 x.D1Nfori,
                                 x.D1Seriori,
                                 x.D1Datori
                             })
                             .Select(x => new RelatorioDevolucaoFat
                             {
                                 Total = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc),
                                 TotalBrut = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valipi),
                             })
                             .ToList();
                #endregion

                Faturamento.Add(new ValoresEmAberto { Nome = "INTERMEDIC", Valor = resultadoInter.Sum(x => x.Total) - DevolInter.Sum(x => x.Total) });

                #endregion

                #region Denuo

                #region Faturado
                var queryDenuo = (from SD20 in ProtheusDenuo.Sd2010s
                                  join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                  join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                  join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                  where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                  && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                  (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains(SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                  && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                  select new
                                  {
                                      Login = SA30.A3Xlogin,
                                      NF = SD20.D2Doc,
                                      Total = SD20.D2Total,

                                  });

                var resultadoDenuo = queryDenuo.GroupBy(x => new
                {
                    x.Login,
                    x.NF,
                    x.Total
                }).Select(x => new RelatorioCirurgiasFaturadas
                {
                    A3Nome = x.Key.Login.Trim(),
                    Nf = x.Key.NF,
                    Total = x.Sum(c => c.Total),
                }).ToList();
                #endregion

                #region Devolucao
                var DevolDenuo = (from SD10 in ProtheusDenuo.Sd1010s
                                  join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SA30 in ProtheusDenuo.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                  join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                  join SA10 in ProtheusDenuo.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                  where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                  && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                  && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                  && (int)(object)SD10.D1Dtdigit >= 20200801 && SC50.C5Utpoper == "F"
                                  orderby SA30.A3Nome
                                  select new
                                  {
                                      SD10.D1Filial,
                                      SD10.D1Fornece,
                                      SD10.D1Loja,
                                      SA10.A1Nome,
                                      SA10.A1Clinter,
                                      SD10.D1Doc,
                                      SD10.D1Serie,
                                      SD10.D1Dtdigit,
                                      SD10.D1Total,
                                      SD10.D1Valdesc,
                                      SD10.D1Valipi,
                                      SD10.D1Valicm,
                                      SA30.A3Nome,
                                      SD10.D1Nfori,
                                      SD10.D1Seriori,
                                      SD10.D1Datori,
                                      SD10.D1Emissao
                                  }
                             )
                             .GroupBy(x => new
                             {
                                 x.A1Nome,
                                 x.A3Nome,
                                 x.D1Filial,
                                 x.D1Fornece,
                                 x.D1Loja,
                                 x.A1Clinter,
                                 x.D1Doc,
                                 x.D1Serie,
                                 x.D1Emissao,
                                 x.D1Dtdigit,
                                 x.D1Nfori,
                                 x.D1Seriori,
                                 x.D1Datori
                             })
                             .Select(x => new RelatorioDevolucaoFat
                             {
                                 Total = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc),
                                 TotalBrut = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valipi),
                             })
                             .ToList();
                #endregion

                Faturamento.Add(new ValoresEmAberto { Nome = "DENUO", Valor = resultadoDenuo.Sum(x => x.Total) - DevolDenuo.Sum(x => x.Total) });

                #endregion

                #region SubDistribuidorInter

                var FaturadoSubInter = (from SD10 in ProtheusInter.Sd2010s
                                        join SF10 in ProtheusInter.Sf4010s on SD10.D2Tes equals SF10.F4Codigo
                                        join SB10 in ProtheusInter.Sb1010s on SD10.D2Cod equals SB10.B1Cod
                                        join SC60 in ProtheusInter.Sc6010s on new { Filial = SD10.D2Filial, NumP = SD10.D2Pedido, Item = SD10.D2Itempv } equals new { Filial = SC60.C6Filial, NumP = SC60.C6Num, Item = SC60.C6Item }
                                        join SC50 in ProtheusInter.Sc5010s on new { Filial2 = SC60.C6Filial, NumC = SC60.C6Num } equals new { Filial2 = SC50.C5Filial, NumC = SC50.C5Num }
                                        join SA10 in ProtheusInter.Sa1010s on new { Codigo = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                                        join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                        where SD10.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*"
                                        && SF10.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*"
                                        && SF10.F4Duplic == "S" && SA10.A1Clinter == "S"
                                        && (int)(object)SD10.D2Emissao >= (int)(object)DataInicio && (int)(object)SD10.D2Emissao <= (int)(object)DataFim
                                        orderby SD10.D2Emissao
                                        select new RelatorioSubDistribuidor
                                        {
                                            Total = SD10.D2Total,
                                            Descon = SD10.D2Descon,
                                        }
                            ).ToList();

                Faturamento.Add(new ValoresEmAberto { Nome = "SubDistribuidor Inter", Valor = FaturadoSubInter.Sum(c => c.Total) });

                #endregion

                #region SubDistribuidorDenuo

                var FaturadoSubDenuo = (from SD10 in ProtheusDenuo.Sd2010s
                                        join SF10 in ProtheusDenuo.Sf4010s on SD10.D2Tes equals SF10.F4Codigo
                                        join SB10 in ProtheusDenuo.Sb1010s on SD10.D2Cod equals SB10.B1Cod
                                        join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SD10.D2Filial, NumP = SD10.D2Pedido, Item = SD10.D2Itempv } equals new { Filial = SC60.C6Filial, NumP = SC60.C6Num, Item = SC60.C6Item }
                                        join SC50 in ProtheusDenuo.Sc5010s on new { Filial2 = SC60.C6Filial, NumC = SC60.C6Num } equals new { Filial2 = SC50.C5Filial, NumC = SC50.C5Num }
                                        join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                                        join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                        where SD10.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*"
                                        && SF10.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*"
                                        && SF10.F4Duplic == "S" && SA10.A1Clinter == "S"
                                        && (int)(object)SD10.D2Emissao >= (int)(object)DataInicio && (int)(object)SD10.D2Emissao <= (int)(object)DataFim
                                        orderby SD10.D2Emissao
                                        select new RelatorioSubDistribuidor
                                        {
                                            Total = SD10.D2Total,
                                            Descon = SD10.D2Descon,
                                        }
                            ).ToList();

                Faturamento.Add(new ValoresEmAberto { Nome = "SubDistribuidor Denuo", Valor = FaturadoSubDenuo.Sum(c => c.Total) });

                #endregion

                #region Dental

                #region Faturado
                var queryDental = (from SD20 in ProtheusDenuo.Sd2010s
                                   join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                   join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                   join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                   join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                                   join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num, Cli = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                                   join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod into sr
                                   from c in sr.DefaultIfEmpty()
                                   where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*"
                                   && SC60.DELET != "*" && SC50.DELET != "*" && c.DELET != "*" &&
                                   (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114)
                                   || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114)
                                   || ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114)
                                   || CF.Contains(SD20.D2Cf)) && (int)(object)SD20.D2Emissao >= (int)(object)DataInicio
                                   && (int)(object)SD20.D2Emissao <= (int)(object)DataFim && SD20.D2Quant != 0 && SC50.C5Xtipopv == "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140"
                                   && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                   select new RelatorioFaturamentoDental
                                   {
                                       D2_FILIAL = SD20.D2Filial,
                                       D2_CLIENTE = SD20.D2Cliente,
                                       D2_LOJA = SD20.D2Emissao,
                                       A1_NOME = SA10.A1Nome,
                                       A1_CLINTER = SA10.A1Clinter,
                                       D2_DOC = SD20.D2Doc,
                                       D2_SERIE = SD20.D2Serie,
                                       D2_EMISSAO = SD20.D2Emissao,
                                       D2_PEDIDO = SD20.D2Pedido,
                                       C5_UNUMAGE = SC50.C5Unumage,
                                       C5_EMISSAO = SC50.C5Emissao,
                                       C5_VEND1 = SC50.C5Vend1,
                                       A3_NOME = c.A3Nome,
                                       C5_X_DTCIR = SC50.C5XDtcir,
                                       C5_X_NMMED = SC50.C5XNmmed,
                                       C5_X_NMPAC = SC50.C5XNmpac,
                                       C5_X_NMPLA = SC50.C5XNmpla,
                                       C5_UTPOPER = SC50.C5Utpoper,
                                       C6_PRODUTO = SC60.C6Produto,
                                       B1_DESC = SB10.B1Desc,
                                       D2_TOTAL = SD20.D2Total,
                                       D2_QUANT = SD20.D2Quant,
                                       D2_VALIPI = SD20.D2Valipi,
                                       D2_VALICM = SD20.D2Valicm,
                                       D2_DESCON = SD20.D2Descon
                                   }
                                  ).GroupBy(Prod =>
                                  new
                                  {
                                      Prod.D2_FILIAL,
                                      Prod.D2_CLIENTE,
                                      Prod.D2_LOJA,
                                      Prod.A1_NOME,
                                      Prod.A1_CLINTER,
                                      Prod.D2_DOC,
                                      Prod.D2_SERIE,
                                      Prod.D2_EMISSAO,
                                      Prod.D2_PEDIDO,
                                      Prod.C5_UNUMAGE,
                                      Prod.C5_EMISSAO,
                                      Prod.C5_VEND1,
                                      Prod.A3_NOME,
                                      Prod.C5_X_DTCIR,
                                      Prod.C5_X_NMMED,
                                      Prod.C5_X_NMPAC,
                                      Prod.C5_X_NMPLA,
                                      Prod.C5_UTPOPER,
                                      Prod.C6_PRODUTO,
                                      Prod.B1_DESC
                                  });

                var FatuDental = queryDental.Select(x => new RelatorioFaturamentoDental
                {
                    D2_TOTAL = x.Sum(c => c.D2_TOTAL),
                }).ToList();
                #endregion

                #region Devolucao
                var CfDevolucao = new string[] { "1202", "2202", "3202", "1553", "2553" };

                var queryDevolucaoDental = (from SD10 in ProtheusDenuo.Sd1010s
                                            join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Fornece = SF20.F2Cliente, Loja = SF20.F2Loja }
                                            join SD20 in ProtheusDenuo.Sd2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja, Item = SD10.D1Itemori } equals new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Fornece = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Item }
                                            join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                            join SA10 in ProtheusDenuo.Sa1010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA10.A1Cod, Loja = SA10.A1Loja }
                                            join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                            where SD10.DELET != "*" && SF20.DELET != "*" && SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SC50.DELET != "*"
                                            && CfDevolucao.Contains(SD10.D1Cf) && SC50.C5Xtipopv == "D"
                                            && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio
                                            && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                            select new
                                            {
                                                SD10.D1Filial,
                                                SD10.D1Fornece,
                                                SD10.D1Loja,
                                                SA10.A1Nome,
                                                SA10.A1Clinter,
                                                SD10.D1Doc,
                                                SD10.D1Serie,
                                                SD10.D1Emissao,
                                                SD20.D2Doc,
                                                SD20.D2Emissao,
                                                SD10.D1Total,
                                                SD10.D1Valipi,
                                                SD10.D1Valicm,
                                                SC50.C5Vend1,
                                                SC50.C5Nomvend,
                                                SD10.D1Valdesc,
                                                SD10.D1Dtdigit,
                                                SA10.A1Est,
                                                SA10.A1Mun
                                            })
                                      .GroupBy(x => new
                                      {
                                          x.D1Filial,
                                          x.D1Fornece,
                                          x.D1Loja,
                                          x.A1Nome,
                                          x.A1Clinter,
                                          x.A1Est,
                                          x.A1Mun,
                                          x.D1Doc,
                                          x.D1Serie,
                                          x.D1Emissao,
                                          x.D1Dtdigit,
                                          x.D2Doc,
                                          x.D2Emissao,
                                          x.C5Vend1,
                                          x.C5Nomvend
                                      });

                var DevolucaoDental = queryDevolucaoDental.Select(x => new RelatorioDevolucaoDental
                {
                    Total = -(x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valipi)),
                }).ToList();

                #endregion

                Faturamento.Add(new ValoresEmAberto { Nome = "DENTAL", Valor = FatuDental.Sum(x => x.D2_TOTAL) + DevolucaoDental.Sum(x => x.Total) });
                #endregion

                #endregion

                var valores = new
                {
                    Valores = Faturamento,
                    ValorTotal = Faturamento.Sum(x => x.Valor)
                };

                return new JsonResult(valores);
            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex, SGID, "DashBoardDiretoria2 Faturado", user);
                return new JsonResult("");
            }
        }

        public JsonResult OnPostEmAberto(string Tipo)
        {
            try
            {
                #region Parametros
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                DateTime data;
                string mesInicio = "";
                string anoInicio = "";
                string DataInicio = $"";

                DateTime Fim;

                string mesFim = "";
                string anoFim = "";
                string DataFim = $"";

                string[] CF = new string[] { "5551", "6551", "6107", "6109" };

                switch (Tipo)
                {
                    case "90+":
                        {
                            Fim = DateTime.Now;

                            mesFim = Fim.Month.ToString("D2");
                            anoFim = Fim.Year.ToString();
                            DataFim = $"{anoFim}{mesFim}31";

                            break;
                        }
                    case "90":
                        {
                            data = DateTime.Now.AddMonths(10);
                            mesInicio = data.Month.ToString("D2");
                            anoInicio = (data.Year - 1).ToString();
                            DataInicio = $"{anoInicio}{mesInicio}01";

                            Fim = DateTime.Now;

                            mesFim = Fim.Month.ToString("D2");
                            anoFim = Fim.Year.ToString();
                            DataFim = $"{anoFim}{mesFim}31";

                            break;
                        }
                    case "60":
                        {
                            data = DateTime.Now.AddMonths(11);
                            mesInicio = data.Month.ToString("D2");
                            anoInicio = (data.Year - 1).ToString();
                            DataInicio = $"{anoInicio}{mesInicio}01";

                            Fim = DateTime.Now;

                            mesFim = Fim.Month.ToString("D2");
                            anoFim = Fim.Year.ToString();
                            DataFim = $"{anoFim}{mesFim}31";

                            break;
                        }
                    default:
                        {
                            data = DateTime.Now;
                            mesInicio = data.Month.ToString("D2");
                            anoInicio = data.Year.ToString();
                            DataInicio = $"{anoInicio}{mesInicio}01";

                            Fim = DateTime.Now;

                            mesFim = Fim.Month.ToString("D2");
                            anoFim = Fim.Year.ToString();
                            DataFim = $"{anoFim}{mesFim}31"; ;

                            break;
                        }
                }
                #endregion

                if (Tipo == "90+")
                {
                    #region EmAberto

                    #region Intermedic
                    var resultadoEmAbertoInter = (from SC5 in ProtheusInter.Sc5010s
                                                  from SC6 in ProtheusInter.Sc6010s
                                                  from SA1 in ProtheusInter.Sa1010s
                                                  from SA3 in ProtheusInter.Sa3010s
                                                  from SF4 in ProtheusInter.Sf4010s
                                                  from SB1 in ProtheusInter.Sb1010s
                                                  where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                                                  && SC6.C6Nota == "" && SC5.C5Nota == "" && SC6.C6Blq != "R" && SC6.DELET != "*"
                                                  && SF4.F4Codigo == SC6.C6Tes && SF4.F4Duplic == "S" && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                                                  && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                                                  && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T") && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                                                  && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                                                  && (int)(object)SC5.C5Emissao <= (int)(object)DataFim
                                                  orderby SC5.C5Num, SC5.C5Emissao descending
                                                  select SC6.C6Valor
                                             ).Sum();

                    ValoresEmAberto.Add(new ValoresEmAberto { Nome = "INTERMEDIC", Valor = resultadoEmAbertoInter });
                    #endregion

                    #region Denuo
                    var resultadoEmAbertoDenuo = (from SC5 in ProtheusDenuo.Sc5010s
                                                  from SC6 in ProtheusDenuo.Sc6010s
                                                  from SA1 in ProtheusDenuo.Sa1010s
                                                  from SA3 in ProtheusDenuo.Sa3010s
                                                  from SF4 in ProtheusDenuo.Sf4010s
                                                  from SB1 in ProtheusDenuo.Sb1010s
                                                  where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                                                  && SC6.C6Nota == "" && SC5.C5Nota == "" && SC6.C6Blq != "R" && SC6.DELET != "*"
                                                  && SF4.F4Codigo == SC6.C6Tes && SF4.F4Duplic == "S" && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                                                  && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                                                  && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T") && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                                                  && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                                                  && SC5.C5Xtipopv != "D"
                                                  && (int)(object)SC5.C5Emissao <= (int)(object)DataFim
                                                  orderby SC5.C5Num, SC5.C5Emissao descending
                                                  select SC6.C6Valor)
                                                  .Sum();

                    ValoresEmAberto.Add(new ValoresEmAberto { Nome = "DENUO", Valor = resultadoEmAbertoDenuo });
                    #endregion

                    #region Dental
                    var resultadoEmAbertoDental = (from SC5 in ProtheusDenuo.Sc5010s
                                                   from SC6 in ProtheusDenuo.Sc6010s
                                                   from SA1 in ProtheusDenuo.Sa1010s
                                                   from SA3 in ProtheusDenuo.Sa3010s
                                                   from SF4 in ProtheusDenuo.Sf4010s
                                                   from SB1 in ProtheusDenuo.Sb1010s
                                                   where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                                                   && SC6.C6Nota == "" && SC6.C6Blq != "R" && SC6.DELET != "*"
                                                   && SF4.F4Codigo == SC6.C6Tes && SF4.F4Duplic == "S" && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                                                   && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                                                   && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T") && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                                                   && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                                                   && SC5.C5Xtipopv == "D"
                                                   && (int)(object)SC5.C5Emissao <= (int)(object)DataFim
                                                   orderby SC5.C5Num, SC5.C5Emissao descending
                                                   select SC6.C6Valor
                                                  ).Sum();

                    ValoresEmAberto.Add(new ValoresEmAberto { Nome = "DENTAL", Valor = resultadoEmAbertoDental });
                    #endregion

                    #region Sub Inter
                    var SubdistribuidorInter = (from SC50 in ProtheusInter.Sc5010s
                                                join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                where SC50.DELET != "*" && SA10.DELET != "*" && SC60.DELET != "*" &&
                                                SA30.DELET != "*" && SA10.A1Clinter == "S" && SC50.C5Nota == "" &&
                                                SC60.C6Qtdven - SC60.C6Qtdent != 0
                                                && SA10.A1Cgc.Substring(0, 8) != "04715053"
                                                && (int)(object)SC50.C5Emissao <= (int)(object)DataFim
                                                orderby SA10.A1Nome, SC50.C5Emissao
                                                select (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven
                                 ).Sum();

                    ValoresEmAberto.Add(new ValoresEmAberto { Nome = "Subdistribuidor Inter", Valor = SubdistribuidorInter });
                    #endregion

                    #region Sub Denuo
                    var SubdistribuidorDenuo = (from SC50 in ProtheusDenuo.Sc5010s
                                                join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                where SC50.DELET != "*" && SA10.DELET != "*" && SC60.DELET != "*" &&
                                                SA30.DELET != "*" && SA10.A1Clinter == "S" && SC50.C5Nota == "" &&
                                                SC60.C6Qtdven - SC60.C6Qtdent != 0
                                                && (int)(object)SC50.C5Emissao <= (int)(object)DataFim
                                                orderby SA10.A1Nome, SC50.C5Emissao
                                                select (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven
                                 ).Sum();

                    ValoresEmAberto.Add(new ValoresEmAberto { Nome = "Subdistribuidor Denuo", Valor = SubdistribuidorDenuo });
                    #endregion

                    #endregion
                }
                else
                {
                    #region EmAberto

                    #region Intermedic
                    var resultadoEmAbertoInter = (from SC5 in ProtheusInter.Sc5010s
                                                  from SC6 in ProtheusInter.Sc6010s
                                                  from SA1 in ProtheusInter.Sa1010s
                                                  from SA3 in ProtheusInter.Sa3010s
                                                  from SF4 in ProtheusInter.Sf4010s
                                                  from SB1 in ProtheusInter.Sb1010s
                                                  where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                                                  && SC6.C6Nota == "" && SC5.C5Nota == "" && SC6.C6Blq != "R" && SC6.DELET != "*"
                                                  && SF4.F4Codigo == SC6.C6Tes && SF4.F4Duplic == "S" && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                                                  && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                                                  && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T") && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                                                  && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                                                  && (int)(object)SC5.C5Emissao >= (int)(object)DataInicio
                                                  && (int)(object)SC5.C5Emissao <= (int)(object)DataFim
                                                  orderby SC5.C5Num, SC5.C5Emissao descending
                                                  select SC6.C6Valor
                                             ).Sum();

                    ValoresEmAberto.Add(new ValoresEmAberto { Nome = "INTERMEDIC", Valor = resultadoEmAbertoInter });
                    #endregion

                    #region Denuo
                    var resultadoEmAbertoDenuo = (from SC5 in ProtheusDenuo.Sc5010s
                                                  from SC6 in ProtheusDenuo.Sc6010s
                                                  from SA1 in ProtheusDenuo.Sa1010s
                                                  from SA3 in ProtheusDenuo.Sa3010s
                                                  from SF4 in ProtheusDenuo.Sf4010s
                                                  from SB1 in ProtheusDenuo.Sb1010s
                                                  where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                                                  && SC6.C6Nota == "" && SC5.C5Nota == "" && SC6.C6Blq != "R" && SC6.DELET != "*"
                                                  && SF4.F4Codigo == SC6.C6Tes && SF4.F4Duplic == "S" && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                                                  && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                                                  && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T") && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                                                  && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                                                  && SC5.C5Xtipopv != "D"
                                                  && (int)(object)SC5.C5Emissao >= (int)(object)DataInicio
                                                  && (int)(object)SC5.C5Emissao <= (int)(object)DataFim
                                                  orderby SC5.C5Num, SC5.C5Emissao descending
                                                  select SC6.C6Valor)
                                                  .Sum();

                    ValoresEmAberto.Add(new ValoresEmAberto { Nome = "DENUO", Valor = resultadoEmAbertoDenuo });
                    #endregion

                    #region Dental
                    var resultadoEmAbertoDental = (from SC5 in ProtheusDenuo.Sc5010s
                                                   from SC6 in ProtheusDenuo.Sc6010s
                                                   from SA1 in ProtheusDenuo.Sa1010s
                                                   from SA3 in ProtheusDenuo.Sa3010s
                                                   from SF4 in ProtheusDenuo.Sf4010s
                                                   from SB1 in ProtheusDenuo.Sb1010s
                                                   where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                                                   && SC6.C6Nota == "" && SC6.C6Blq != "R" && SC6.DELET != "*"
                                                   && SF4.F4Codigo == SC6.C6Tes && SF4.F4Duplic == "S" && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                                                   && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                                                   && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T") && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                                                   && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                                                   && SC5.C5Xtipopv == "D"
                                                   && (int)(object)SC5.C5Emissao >= (int)(object)DataInicio
                                                   && (int)(object)SC5.C5Emissao <= (int)(object)DataFim
                                                   orderby SC5.C5Num, SC5.C5Emissao descending
                                                   select SC6.C6Valor
                                                  ).Sum();

                    ValoresEmAberto.Add(new ValoresEmAberto { Nome = "DENTAL", Valor = resultadoEmAbertoDental });
                    #endregion

                    #region Sub Inter
                    var SubdistribuidorInter = (from SC50 in ProtheusInter.Sc5010s
                                                join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                where SC50.DELET != "*" && SA10.DELET != "*" && SC60.DELET != "*" &&
                                                SA30.DELET != "*" && SA10.A1Clinter == "S" && SC50.C5Nota == "" &&
                                                SC60.C6Qtdven - SC60.C6Qtdent != 0
                                                && SA10.A1Cgc.Substring(0, 8) != "04715053"
                                                && (int)(object)SC50.C5Emissao >= (int)(object)DataInicio
                                                && (int)(object)SC50.C5Emissao <= (int)(object)DataFim
                                                orderby SA10.A1Nome, SC50.C5Emissao
                                                select (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven
                                 ).Sum();

                    ValoresEmAberto.Add(new ValoresEmAberto { Nome = "Subdistribuidor Inter", Valor = SubdistribuidorInter });
                    #endregion

                    #region Sub Denuo
                    var SubdistribuidorDenuo = (from SC50 in ProtheusDenuo.Sc5010s
                                                join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                where SC50.DELET != "*" && SA10.DELET != "*" && SC60.DELET != "*" &&
                                                SA30.DELET != "*" && SA10.A1Clinter == "S" && SC50.C5Nota == "" &&
                                                SC60.C6Qtdven - SC60.C6Qtdent != 0
                                                && (int)(object)SC50.C5Emissao >= (int)(object)DataInicio
                                                && (int)(object)SC50.C5Emissao <= (int)(object)DataFim
                                                orderby SA10.A1Nome, SC50.C5Emissao
                                                select (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven
                                 ).Sum();

                    ValoresEmAberto.Add(new ValoresEmAberto { Nome = "Subdistribuidor Denuo", Valor = SubdistribuidorDenuo });
                    #endregion

                    #endregion

                }


                var valores = new
                {
                    Valores = ValoresEmAberto,
                    ValorTotal = ValoresEmAberto.Sum(x => x.Valor)
                };

                return new JsonResult(valores);
            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex, SGID, "DashBoardDiretoria2", user);

                return new JsonResult("");
            }
        }

        public JsonResult OnPostBaixa(string Tipo)
        {
            try
            {
                #region Parametros
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                DateTime data;
                string mesInicio = "";
                string anoInicio = "";
                string DataInicio = $"";

                DateTime Fim;

                string mesFim = "";
                string anoFim = "";
                string DataFim = $"";

                string[] CF = new string[] { "5551", "6551", "6107", "6109" };

                switch (Tipo)
                {
                    case "30":
                        {
                            data = DateTime.Now;
                            mesInicio = data.Month.ToString("D2");
                            anoInicio = data.Year.ToString();
                            DataInicio = $"{anoInicio}{mesInicio}01";

                            DataFim = $"{anoInicio}{mesInicio}31";

                            break;

                            
                        }
                    case "60":
                        {
                            data = DateTime.Now.AddMonths(11);
                            mesInicio = data.Month.ToString("D2");
                            anoInicio = (data.Year - 1).ToString();
                            DataInicio = $"{anoInicio}{mesInicio}01";

                            Fim = DateTime.Now;

                            mesFim = Fim.Month.ToString("D2");
                            anoFim = Fim.Year.ToString();
                            DataFim = $"{anoFim}{mesFim}31";

                            break;
                        }
                    default:
                        {
                            data = DateTime.Now.AddMonths(10);
                            mesInicio = data.Month.ToString("D2");
                            anoInicio = (data.Year - 1).ToString();
                            DataInicio = $"{anoInicio}{mesInicio}01";

                            Fim = DateTime.Now;

                            mesFim = Fim.Month.ToString("D2");
                            anoFim = Fim.Year.ToString();
                            DataFim = $"{anoFim}{mesFim}31";

                            break;

                        }
                }
                #endregion

                #region Baixa(Caixa)

                #region Intermedic

                var BaixaInter = (from SE50 in ProtheusInter.Se5010s
                                  join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                  equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                  join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                  join SC50 in ProtheusInter.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                  where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                  && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                  && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                  && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                  && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                  && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                  && SA10.A1Clinter != "S"
                                  select new RelatorioAreceberBaixa
                                  {
                                      TotalBaixado = SE50.E5Valor,
                                  }
                             ).ToList();

                Baixa.Add(new ValoresEmAberto { Nome = "INTERMEDIC", Valor = BaixaInter.Sum(x => x.TotalBaixado) });

                #endregion

                #region Denuo
                var BaixaDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                  join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                  equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                  join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                  join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                  where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                  && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                  && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                  && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                  && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                  && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                  && SA10.A1Clinter != "S" && SC50.C5Xtipopv != "D"
                                  select new RelatorioAreceberBaixa
                                  {
                                      TotalBaixado = SE50.E5Valor,
                                  }
                             ).ToList();

                Baixa.Add(new ValoresEmAberto { Nome = "DENUO", Valor = BaixaDenuo.Sum(x => x.TotalBaixado) });
                #endregion

                #region Sub Inter

                var BaixaSubInter = (from SE50 in ProtheusInter.Se5010s
                                     join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                     equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                     join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                     join SC50 in ProtheusInter.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                     where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                     && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                     && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                     && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                     && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                     && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                     && SA10.A1Clinter == "S"
                                     select new RelatorioAreceberBaixa
                                     {

                                         TotalBaixado = SE50.E5Valor,
                                     }
                             ).ToList();

                Baixa.Add(new ValoresEmAberto { Nome = "SubDistribuidor Inter", Valor = BaixaSubInter.Sum(x => x.TotalBaixado) });

                #endregion

                #region Sub Denuo

                var BaixaSubDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                     join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                     equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                     join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                     join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                     where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                     && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                     && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                     && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                     && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                     && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                     && SA10.A1Clinter == "S"
                                     select new RelatorioAreceberBaixa
                                     {
                                         TotalBaixado = SE50.E5Valor,
                                     }
                             ).ToList();

                Baixa.Add(new ValoresEmAberto { Nome = "SubDistribuidor Denuo", Valor = BaixaSubDenuo.Sum(x => x.TotalBaixado) });


                #endregion

                #region Dental

                var BaixaDenuoDental = (from SE50 in ProtheusDenuo.Se5010s
                                        join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                        equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                        join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                        join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                        where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                        && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                        && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                        && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                        && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                        && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                        && SA10.A1Clinter != "S" && SC50.C5Xtipopv == "D"
                                        select new RelatorioAreceberBaixa
                                        {
                                            TotalBaixado = SE50.E5Valor,
                                        }
                            ).ToList();

                Baixa.Add(new ValoresEmAberto { Nome = "Dental", Valor = BaixaDenuoDental.Sum(x => x.TotalBaixado) });

                #endregion

                #endregion

                var valores = new
                {
                    Valores = Baixa,
                    ValorTotal = Baixa.Sum(x => x.Valor)
                };

                return new JsonResult(valores);

            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex, SGID, "DashBoardDiretoria2", user);
                return new JsonResult("");
            }
        }

        public JsonResult OnPostVencidos()
        {
            string user = User.Identity.Name.Split("@")[0].ToUpper();
            DateTime data = DateTime.Now.AddMonths(9);
            string mesInicio = data.Month.ToString("D2");
            string anoInicio = (data.Year - 1).ToString();
            string DataInicio = $"{data.Year - 1}{data.Month.ToString("D2")}{data.Day.ToString("D2")}";

            var NaoBaixados = (from SE10 in ProtheusDenuo.Se1010s  
                               join SA10 in ProtheusDenuo.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                               where  SE10.DELET != "*" 
                               && (int)(object)SE10.E1Vencrea <= (int)(object)DataInicio
                               && SA10.A1Clinter != "G" && SA10.A1Msblql != "1" && SE10.E1Baixa == ""
                               select new TitulosVencidos
                               {
                                   CodCliente = SE10.E1Cliente,
                                   Cliente = SE10.E1Nomcli,
                                   Valor = SE10.E1Valor
                               }
                            )
                            .GroupBy(x=> new
                            {
                                x.CodCliente,
                                x.Cliente,
                            })
                            .Select(x=> new TitulosVencidos
                            {
                                CodCliente = x.Key.CodCliente,
                                Cliente = x.Key.Cliente,
                                Valor = x.Sum(c=>c.Valor)
                            }).OrderByDescending(x => x.Valor).Take(5).ToList();

                var valores = new
                {
                    Valores = NaoBaixados,
                    ValorTotal = NaoBaixados.Sum(x => x.Valor)
                };

            return new JsonResult(valores);
        }

        public JsonResult OnPostVencidosInter()
        {
            string user = User.Identity.Name.Split("@")[0].ToUpper();
            DateTime data = DateTime.Now.AddMonths(9);
            string mesInicio = data.Month.ToString("D2");
            string anoInicio = (data.Year - 1).ToString();
            string DataInicio = $"{data.Year - 1}{data.Month.ToString("D2")}{data.Day.ToString("D2")}";

            var NaoBaixados = (from SE10 in ProtheusInter.Se1010s
                               join SA10 in ProtheusInter.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                               where SE10.DELET != "*"
                               && (int)(object)SE10.E1Vencrea <= (int)(object)DataInicio
                               && SA10.A1Clinter != "G" && SA10.A1Msblql != "1" && SE10.E1Baixa == ""
                               select new TitulosVencidos
                               {
                                   CodCliente = SE10.E1Cliente,
                                   Cliente = SE10.E1Nomcli,
                                   Valor = SE10.E1Valor
                               }
                            )
                            .GroupBy(x => new
                            {
                                x.CodCliente,
                                x.Cliente,
                            })
                            .Select(x => new TitulosVencidos
                            {
                                CodCliente = x.Key.CodCliente,
                                Cliente = x.Key.Cliente,
                                Valor = x.Sum(c => c.Valor)
                            }).OrderByDescending(x => x.Valor).Take(5).ToList();

            var valores = new
            {
                Valores = NaoBaixados,
                ValorTotal = NaoBaixados.Sum(x => x.Valor)
            };

            return new JsonResult(valores);
        }

        public JsonResult OnPostFaturamentoPorLinha(string Tipo)
        {
            var LinhasValor = new List<RelatorioFaturamentoLinhas>();

            #region Parametros
            string user = User.Identity.Name.Split("@")[0].ToUpper();
            DateTime data;
            string mesInicio = "";
            string anoInicio = "";
            string DataInicio = $"";

            DateTime Fim;

            string mesFim = "";
            string anoFim = "";
            string DataFim = $"";

            string[] CF = new string[] { "5551", "6551", "6107", "6109" };

            switch (Tipo)
            {
                case "30":
                    {
                        data = DateTime.Now;
                        mesInicio = data.Month.ToString("D2");
                        anoInicio = data.Year.ToString();
                        DataInicio = $"{anoInicio}{mesInicio}01";

                        DataFim = $"{anoInicio}{mesInicio}31";

                        break;

                    }
                case "60":
                    {
                        data = DateTime.Now.AddMonths(11);
                        mesInicio = data.Month.ToString("D2");
                        anoInicio = (data.Year - 1).ToString();
                        DataInicio = $"{anoInicio}{mesInicio}01";

                        Fim = DateTime.Now;

                        mesFim = Fim.Month.ToString("D2");
                        anoFim = Fim.Year.ToString();
                        DataFim = $"{anoFim}{mesFim}31";

                        break;
                    }
                default:
                    {

                        data = DateTime.Now.AddMonths(10);
                        mesInicio = data.Month.ToString("D2");
                        anoInicio = (data.Year - 1).ToString();
                        DataInicio = $"{anoInicio}{mesInicio}01";

                        Fim = DateTime.Now;

                        mesFim = Fim.Month.ToString("D2");
                        anoFim = Fim.Year.ToString();
                        DataFim = $"{anoFim}{mesFim}31";

                        break;
                    }
            }
            #endregion

            var linhasInter = ProtheusInter.Sa3010s.Where(x => x.A3Xdescun != "").Select(x => x.A3Xdescun.Trim()).Distinct().ToList();
            var linhasDenuo = ProtheusDenuo.Sa3010s.Where(x => x.A3Xdescun != "").Select(x => x.A3Xdescun.Trim()).Distinct().ToList();

            var Linhas = linhasInter.Union(linhasDenuo).ToList();


            Linhas.ForEach(x =>
            {
                LinhasValor.Add(new RelatorioFaturamentoLinhas { Nome = x });
            });


            #region Faturado Inter
            var query = (from SD20 in ProtheusInter.Sd2010s
                         join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                         join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                         join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                         join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains(SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                         && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                         select new
                         {
                             Login = SA30.A3Xlogin,
                             NF = SD20.D2Doc,
                             Total = SD20.D2Total,
                             Linha = SA30.A3Xdescun

                         });

            var resultadoInter = query.GroupBy(x => new
            {
                x.Login,
                x.NF,
                x.Linha
            }).Select(x => new RelatorioCirurgiasFaturadas
            {
                A3Nome = x.Key.Login.Trim(),
                Nf = x.Key.NF,
                Total = x.Sum(c => c.Total),
                Linha = x.Key.Linha
            }).ToList();
            #endregion

            #region Faturado Denuo
            var queryDenuo = (from SD20 in ProtheusDenuo.Sd2010s
                              join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                              join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                              join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                              join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                              join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                              where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                              && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                              (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains(SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                              && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                              select new
                              {
                                  Login = SA30.A3Xlogin,
                                  NF = SD20.D2Doc,
                                  Total = SD20.D2Total,
                                  Linha = SA30.A3Xdescun
                              });

            var resultadoDenuo = queryDenuo.GroupBy(x => new
            {
                x.Login,
                x.NF,
                x.Linha
            }).Select(x => new RelatorioCirurgiasFaturadas
            {
                A3Nome = x.Key.Login.Trim(),
                Nf = x.Key.NF,
                Total = x.Sum(c => c.Total),
                Linha = x.Key.Linha
            }).ToList();
            #endregion

            #region SubDistribuidorInter

            var FaturadoSubInter = (from SD10 in ProtheusInter.Sd2010s
                                    join SF10 in ProtheusInter.Sf4010s on SD10.D2Tes equals SF10.F4Codigo
                                    join SB10 in ProtheusInter.Sb1010s on SD10.D2Cod equals SB10.B1Cod
                                    join SC60 in ProtheusInter.Sc6010s on new { Filial = SD10.D2Filial, NumP = SD10.D2Pedido, Item = SD10.D2Itempv } equals new { Filial = SC60.C6Filial, NumP = SC60.C6Num, Item = SC60.C6Item }
                                    join SC50 in ProtheusInter.Sc5010s on new { Filial2 = SC60.C6Filial, NumC = SC60.C6Num } equals new { Filial2 = SC50.C5Filial, NumC = SC50.C5Num }
                                    join SA10 in ProtheusInter.Sa1010s on new { Codigo = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                                    join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                    where SD10.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*"
                                    && SF10.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*"
                                    && SF10.F4Duplic == "S" && SA10.A1Clinter == "S"
                                    && (int)(object)SD10.D2Emissao >= (int)(object)DataInicio && (int)(object)SD10.D2Emissao <= (int)(object)DataFim
                                    orderby SD10.D2Emissao
                                    select new
                                    {
                                        NF = SD10.D2Doc,
                                        Total = SD10.D2Total,
                                        Linha = SA30.A3Xdescun
                                    }
                                    ).GroupBy(x => new
                                    {
                                        x.NF,
                                        x.Linha
                                    }).Select(x => new
                                    {
                                        Linha = x.Key.Linha,
                                        Quant = x.Count(),
                                        Total = x.Sum(c => c.Total)
                                    })
                                    .ToList();
        
            #endregion

            #region SubDistribuidorDenuo

            var FaturadoSubDenuo = (from SD10 in ProtheusDenuo.Sd2010s
                                    join SF10 in ProtheusDenuo.Sf4010s on SD10.D2Tes equals SF10.F4Codigo
                                    join SB10 in ProtheusDenuo.Sb1010s on SD10.D2Cod equals SB10.B1Cod
                                    join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SD10.D2Filial, NumP = SD10.D2Pedido, Item = SD10.D2Itempv } equals new { Filial = SC60.C6Filial, NumP = SC60.C6Num, Item = SC60.C6Item }
                                    join SC50 in ProtheusDenuo.Sc5010s on new { Filial2 = SC60.C6Filial, NumC = SC60.C6Num } equals new { Filial2 = SC50.C5Filial, NumC = SC50.C5Num }
                                    join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                                    join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                    where SD10.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*"
                                    && SF10.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*"
                                    && SF10.F4Duplic == "S" && SA10.A1Clinter == "S"
                                    && (int)(object)SD10.D2Emissao >= (int)(object)DataInicio && (int)(object)SD10.D2Emissao <= (int)(object)DataFim
                                    orderby SD10.D2Emissao
                                    select new 
                                    {
                                        NF = SD10.D2Doc,
                                        Total = SD10.D2Total,
                                        Linha = SA30.A3Xdescun
                                    }
                                    ).GroupBy(x => new
                                    {
                                        x.NF,
                                        x.Linha
                                    }).Select(x=> new
                                    {
                                        Linha = x.Key.Linha,
                                        Quant = x.Count(),
                                        Total = x.Sum(c=> c.Total)
                                    })
                                    .ToList();

            #endregion

            #region Dental

            var queryDental = (from SD20 in ProtheusDenuo.Sd2010s
                               join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                               join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                               join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                               join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                               join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num, Cli = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                               join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod into sr
                               from c in sr.DefaultIfEmpty()
                               where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*"
                               && SC60.DELET != "*" && SC50.DELET != "*" && c.DELET != "*" &&
                               (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114)
                               || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114)
                               || ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114)
                               || CF.Contains(SD20.D2Cf)) && (int)(object)SD20.D2Emissao >= (int)(object)DataInicio
                               && (int)(object)SD20.D2Emissao <= (int)(object)DataFim && SD20.D2Quant != 0 && SC50.C5Xtipopv == "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140"
                               && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                               select new
                               {
                                   NF = SD20.D2Doc,
                                   TOTAL = SD20.D2Total,
                                   Linha = "DENTAL"
                               }
                                ).GroupBy(Prod =>
                                new
                                {
                                    Prod.NF,
                                    Prod.Linha
                                   
                                }).Select(x => new 
                                {
                                    Linha = x.Key.Linha,
                                    Quant = x.Count(),
                                    TOTAL = x.Sum(c => c.TOTAL),
                                }).ToList();

            #endregion


            var resultado = resultadoDenuo.Concat(resultadoInter).ToList();


            var resultado2 = resultado.GroupBy(x => new { x.Linha }).Select(x => new 
            { 
                Linha = x.Key.Linha,
                Quant = x.Count(),
                Valor = x.Sum(c => c.Total)
            }).ToList();


            var resultado3 = FaturadoSubDenuo.Concat(FaturadoSubInter).ToList();

            LinhasValor.ForEach(x =>
            {
                var valor = resultado2.Where(c => c.Linha.Trim() == x.Nome).FirstOrDefault();

                if (valor != null)
                {
                    x.Quant = valor.Quant;
                    x.Faturamento = valor.Valor;
                }
            });

            LinhasValor.ForEach(x =>
            {
                var valor = queryDental.Where(c => c.Linha.Trim() == x.Nome).ToList();

                if (valor != null || valor.Count > 0)
                {
                    valor.ForEach(c =>
                    {
                        x.Quant += c.Quant;
                        x.Faturamento += c.TOTAL;
                    });

                }
            });

            LinhasValor.ForEach(x =>
            {
                var valor = resultado3.Where(c => c.Linha.Trim() == x.Nome).ToList();

                if (valor != null || valor.Count > 0)
                {
                    valor.ForEach(c =>
                    {
                        x.Quant += c.Quant;
                        x.Faturamento += c.Total;
                    });
                    
                }
            });




            return new JsonResult(LinhasValor);
        }

    }
}