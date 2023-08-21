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
                string[] CF = new string[] {" 5551", "6551", "6107", "6109", "5117", "6117" };
                string DataInicio = $"{Ano}0101";

                string DataFim = $"{Ano}1231";
                if (Mes == "13")
                {
                    #region Parametros

                    DataInicio = $"{Ano}0101";

                    DataFim = $"{Ano}1231";
                    #endregion
                }
                else
                {
                    #region Parametros
                    DataInicio = $"{Ano}{Mes}01";

                    DataFim = $"{Ano}{Mes}31";

                    #endregion
                }

                #region Boxs
                var Linhas = new List<string> { "BUCOMAXILO", "NEURO", "ORTOPEDIA", "TORAX" };


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
                             (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains(SD20.D2Cf)) 
                             && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio 
                             && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                             //&& SD20.D2Quant != 0 
                             && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
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
                    A3Nome = x.Key.Login,
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
                                  (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains(SD20.D2Cf)) 
                                  && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio 
                                  && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                  //&& SD20.D2Quant != 0 
                                  && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
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
                    A3Nome = x.Key.Login,
                    Nf = x.Key.NF,
                    Total = x.Sum(c => c.Total),
                    Linha = x.Key.Linha
                }).ToList();
                #endregion

                var resultado = resultadoDenuo.Concat(resultadoInter).ToList();


                var resultado2 = resultado.GroupBy(x => new { x.Linha }).Select(x => new
                {
                    Linha = x.Key.Linha,
                    Quant = x.Count(),
                    Valor = x.Sum(c => c.Total)
                }).ToList();


                LinhasValor.ForEach(x =>
                {
                    var valor = resultado2.Where(c => c.Linha.Trim() == x.Nome).FirstOrDefault();

                    if (valor != null)
                    {
                        x.Quant = valor.Quant;
                        x.Faturamento = valor.Valor;
                    }

                });


                var LinhasTela = new List<RelatorioFaturamentoLinhas>
                    {
                        new RelatorioFaturamentoLinhas { Nome = "QTDA. CIRURGIAS", Quant = LinhasValor.Sum(x => x.Quant) }
                    };

                LinhasValor.ForEach(x =>
                {
                    LinhasTela.Add(new RelatorioFaturamentoLinhas { Nome = x.Nome, Quant = x.Quant });
                });

                LinhasTela.Add(new RelatorioFaturamentoLinhas { Nome = "VALOR CIRURGIAS", Quant = LinhasValor.Sum(x => x.Faturamento) });


                #endregion

                #region Faturados

                #region Intermedic

                #region Faturado
                var query2 = (from SD20 in ProtheusInter.Sd2010s
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

                var resultadoInter2 = query2.GroupBy(x => new
                {
                    x.Login,
                    x.NF,
                    x.Total
                }).Select(x => new RelatorioCirurgiasFaturadas
                {
                    A3Nome = x.Key.Login,
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

                var resultadoDenuo2 = queryDenuo2.GroupBy(x => new
                {
                    x.Login,
                    x.NF,
                    x.Total
                }).Select(x => new RelatorioCirurgiasFaturadas
                {
                    A3Nome = x.Key.Login,
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

                Faturamento.Add(new ValoresEmAberto { Nome = "DENUO", Valor = resultadoDenuo2.Sum(x => x.Total) - DevolDenuo.Sum(x => x.Total) });

                #endregion


                #endregion

                var valores = new
                {
                    Valores = Faturamento,
                    ValorTotal = Faturamento.Sum(x => x.Valor)
                };


                var Teste = new
                {
                    valores = valores,
                    Linhas = LinhasTela
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
    }
}
