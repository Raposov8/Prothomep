using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Models;
using SGID.Models.Account.RH;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.DashBoards
{
    [Authorize]
    public class DashBoardMetasModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        private TOTVSDENUOContext ProtheusDenuo { get; set; }

        private TOTVSINTERContext ProtheusInter { get; set; }

        public string Ano { get; set; } = "2023";

        public DashBoardMetasModel(ApplicationDbContext sgid,TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            SGID = sgid;
            ProtheusDenuo = denuo;
            ProtheusInter = inter;

        }

        public void OnGet()
        {
            Ano = DateTime.Now.Year.ToString();
        }


        public JsonResult OnPostFaturados(string Mes,string Ano)
        {

            if (Mes == "13")
            {
                var LinhasValor = new List<RelatorioFaturamentoLinhas>();

                #region Parametros
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                string DataInicio = $"{Ano}0101";

                string DataFim = $"{Ano}1231";

                string[] CF = new string[] { "5551", "6551", "6107", "6109" };

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
                    A3Nome = x.Key.Login,
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
                                        }).Select(x => new
                                        {
                                            Linha = x.Key.Linha,
                                            Quant = x.Count(),
                                            Total = x.Sum(c => c.Total)
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
            else
            {
                var LinhasValor = new List<RelatorioFaturamentoLinhas>();

                #region Parametros
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                string DataInicio = $"{Ano}{Mes}01";

                string DataFim = $"{Ano}{Mes}31";

                string[] CF = new string[] { "5551", "6551", "6107", "6109" };

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
                    A3Nome = x.Key.Login,
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
                                        }).Select(x => new
                                        {
                                            Linha = x.Key.Linha,
                                            Quant = x.Count(),
                                            Total = x.Sum(c => c.Total)
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
}
