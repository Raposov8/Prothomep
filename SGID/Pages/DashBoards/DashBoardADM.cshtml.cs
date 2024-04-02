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

namespace SGID.Pages.DashBoards
{
    [Authorize]
    public class DashBoardADMModel : PageModel
    {
        public double FaturadoMesDenuoValor { get; set; }
        public double FaturadoMesInterValor { get; set; }
        public double MetaInter { get; set; } = 2600000;
        public double MetaDenuo { get; set; } = 2600000;
        public double Comissao { get; set; }

        public string Ano { get; set; }

        public TOTVSDENUOContext ProtheusDenuo { get; set; }
        public TOTVSINTERContext ProtheusInter { get; set; }

        public ApplicationDbContext SGID { get; set; }

        public List<ValoresEmAberto> Faturamento { get; set; } = new List<ValoresEmAberto>();

        public List<ValoresEmAberto> ValoresEmAberto { get; set; } = new List<ValoresEmAberto>();

        public DashBoardADMModel(TOTVSDENUOContext denuo, TOTVSINTERContext inter, ApplicationDbContext sgid)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
            SGID = sgid;
        }

        public void OnGet()
        {
            Ano = DateTime.Now.Year.ToString();
        }

        public JsonResult OnPostFaturados(string Mes, string Ano)
        {
            try
            {
                var LinhasValor = new List<RelatorioFaturamentoLinhas>();
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                string[] CF = new string[] { "5551", "6551", "6107", "6109" };
                string DataInicio = $"{Ano}0101";

                string DataFim = $"{Ano}1231";
                if (Mes == "13")
                {
                    #region Parametros

                    DataInicio = $"{Ano}0101";

                    DataFim = $"{Ano}1231";

                    MetaInter = 2600000 * 12;
                    MetaDenuo = 2600000 * 12;
                    #endregion
                }
                else
                {
                    #region Parametros
                    DataInicio = $"{Ano}{Mes}01";

                    DataFim = $"{Ano}{Mes}31";

                    MetaInter = 2600000;
                    MetaDenuo = 2600000;

                    #endregion
                }

                #region Faturados

                #region Intermedic

                #region Faturado
                var query2 = (from SD20 in ProtheusInter.Sd2010s
                              join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                              join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                              join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                              join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                              join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                              where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                              && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                              (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains(SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                              && SD20.D2Quant != 0 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                              select new
                              {
                                  Login = SA30.A3Xlogin,
                                  NF = SD20.D2Doc,
                                  Total = SD20.D2Total,
                                  Linha = SA30.A3Xdescun
                              });

                var resultadoInter2 = query2.GroupBy(x => new
                {
                    x.Login,
                    x.NF,
                    x.Total,
                    x.Linha
                }).Select(x => new RelatorioCirurgiasFaturadas
                {
                    A3Nome = x.Key.Login,
                    Nf = x.Key.NF,
                    Total = x.Sum(c => c.Total),
                    Linha = x.Key.Linha
                }).ToList();
                #endregion

                #region Devolucao
                var DevolInter = (from SD10 in ProtheusInter.Sd1010s
                                  join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                  join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                  join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                  where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
                                  && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                  && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                  && (int)(object)SD10.D1Dtdigit >= 20200801 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K")
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
                                      SD10.D1Emissao,
                                      Linha = SA30.A3Xdescun
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
                                 x.D1Datori,
                                 x.Linha
                             })
                             .Select(x => new RelatorioDevolucaoFat
                             {
                                 Nf = x.Key.D1Doc,
                                 Total = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc),
                                 TotalBrut = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valipi),
                                 Linha = x.Key.Linha
                             })
                             .ToList();
                #endregion

                Faturamento.Add(new ValoresEmAberto { Nome = "INTERMEDIC", Valor = resultadoInter2.Sum(x => x.Total) - DevolInter.Sum(x => x.Total) });

                #endregion

                #region Denuo

                #region Faturado
                var queryDenuo2 = (from SD20 in ProtheusDenuo.Sd2010s
                                   join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                   join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                   join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                   join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                   join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                   where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                   && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                   (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains(SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                   && SD20.D2Quant != 0 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                   select new
                                   {
                                       Login = SA30.A3Xlogin,
                                       NF = SD20.D2Doc,
                                       Total = SD20.D2Total,
                                       Linha = SA30.A3Xdescun

                                   });

                var resultadoDenuo2 = queryDenuo2.GroupBy(x => new
                {
                    x.Login,
                    x.NF,
                    x.Total,
                    x.Linha
                }).Select(x => new RelatorioCirurgiasFaturadas
                {
                    A3Nome = x.Key.Login,
                    Nf = x.Key.NF,
                    Total = x.Sum(c => c.Total),
                    Linha = x.Key.Linha
                }).ToList();
                #endregion

                #region Devolucao
                var DevolDenuo = (from SD10 in ProtheusDenuo.Sd1010s
                                  join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SA30 in ProtheusDenuo.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                  join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                  join SA10 in ProtheusDenuo.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                  where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
                                  && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                  && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                  && (int)(object)SD10.D1Dtdigit >= 20200801 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K")
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
                                      SD10.D1Emissao,
                                      Linha = SA30.A3Xdescun
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
                                 x.D1Datori,
                                 x.Linha
                             })
                             .Select(x => new RelatorioDevolucaoFat
                             {
                                 Nf = x.Key.D1Doc,
                                 Total = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc),
                                 TotalBrut = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valipi),
                                 Linha = x.Key.Linha
                             })
                             .ToList();
                #endregion

                Faturamento.Add(new ValoresEmAberto { Nome = "DENUO", Valor = resultadoDenuo2.Sum(x => x.Total) - DevolDenuo.Sum(x => x.Total) });

                #endregion

                #endregion

                #region Boxs
                var Linhas = new List<string> { "BUCOMAXILO", "NEURO", "ORTOPEDIA", "TORAX" };


                Linhas.ForEach(x =>
                {
                    LinhasValor.Add(new RelatorioFaturamentoLinhas { Nome = x });
                });


                var resultado = resultadoDenuo2.Concat(resultadoInter2).ToList();

                var Devolucao = DevolInter.Concat(DevolDenuo).ToList();

                var resultado2 = resultado.GroupBy(x => new { x.Linha }).Select(x => new
                {
                    Linha = x.Key.Linha,
                    Quant = x.GroupBy(c => c.Nf).Count(),
                    Valor = x.Sum(c => c.Total),
                    ValorBrut = x.Sum(c => c.TotalBrut)
                }).ToList();

                var Devolucao2 = Devolucao.GroupBy(x => new { x.Linha }).Select(x => new
                {
                    Linha = x.Key.Linha,
                    Quant = x.GroupBy(c => c.Nf).Count(),
                    Valor = x.Sum(c => c.Total),
                    ValorBrut = x.Sum(c => c.TotalBrut)

                }).ToList();

                LinhasValor.ForEach(x =>
                {
                    var valor = resultado2.Where(c => c.Linha.Trim() == x.Nome).FirstOrDefault();

                    if (valor != null)
                    {
                        x.Quant = valor.Quant;
                        x.Faturamento = valor.Valor;
                    }

                    var valor2 = Devolucao2.Where(c => c.Linha.Trim() == x.Nome).FirstOrDefault();

                    if (valor2 != null)
                    {
                        x.Quant -= valor2.Quant;
                        x.Faturamento -= valor2.Valor;
                    }

                });





                var LinhasTela = new List<RelatorioFaturamentoLinhas>
                {
                    new RelatorioFaturamentoLinhas { Nome = "QTDA. CIRURGIAS", Quant = resultado2.Sum(x => x.Quant) -  Devolucao2.Sum(x=> x.Quant) }
                };

                LinhasValor.ForEach(x =>
                {
                    LinhasTela.Add(new RelatorioFaturamentoLinhas { Nome = x.Nome, Quant = x.Quant });
                });

                LinhasTela.Add(new RelatorioFaturamentoLinhas { Nome = "INTERMEDIC", Quant = resultadoInter2.Sum(x => x.Total) });

                LinhasTela.Add(new RelatorioFaturamentoLinhas { Nome = "DENUO", Quant = resultadoDenuo2.Sum(x => x.Total) });

                LinhasTela.Add(new RelatorioFaturamentoLinhas { Nome = "VALOR CIRURGIAS", Quant = resultado2.Sum(x => x.Valor) - Devolucao2.Sum(x => x.Valor) });
                #endregion

                var valores = new
                {
                    Valores = Faturamento,
                    ValorTotal = Faturamento.Sum(x => x.Valor)
                };


                var Teste = new
                {
                    valores = valores,
                    Linhas = LinhasTela,
                    Meta = MetaInter + MetaDenuo
                };

                return new JsonResult(Teste);

            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex, SGID, "DashBoardADM Faturados", user);

                return new JsonResult("");
            }
        }

        public JsonResult OnPostEmAberto(string Mes, string Ano)
        {

            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                string[] CF = new string[] { "5551", "6551", "6107", "6109" };
                string DataInicio = "";
                string DataFim = "";

                #region EmAberto

                #region Intermedic
                var resultadoEmAbertoInter = (from SC5 in ProtheusInter.Sc5010s
                                              from SC6 in ProtheusInter.Sc6010s
                                              from SA1 in ProtheusInter.Sa1010s
                                              from SA3 in ProtheusInter.Sa3010s
                                              from SF4 in ProtheusInter.Sf4010s
                                              from SB1 in ProtheusInter.Sb1010s
                                              where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                                              && SC6.C6Nota == "" && SC6.C6Blq != "R" && SC6.DELET != "*"
                                              && SF4.F4Codigo == SC6.C6Tes && SF4.F4Duplic == "S" && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                                              && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                                              && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T") && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                                              && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
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
                                              && SC6.C6Nota == "" && SC6.C6Blq != "R" && SC6.DELET != "*"
                                              && SF4.F4Codigo == SC6.C6Tes && SF4.F4Duplic == "S" && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                                              && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                                              && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T") && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                                              && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                                              && SC5.C5Xtipopv != "D"
                                              orderby SC5.C5Num, SC5.C5Emissao descending
                                              select SC6.C6Valor)
                                              .Sum();

                ValoresEmAberto.Add(new ValoresEmAberto { Nome = "DENUO", Valor = resultadoEmAbertoDenuo });
                #endregion

                #endregion

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
                Logger.Log(ex, SGID, "DashBoardMetas EmAberto", user);

                return new JsonResult("");
            }

        }
    }
}
