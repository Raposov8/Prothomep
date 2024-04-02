using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Comercial;
using SGID.Models.Denuo;

namespace SGID.Pages.DashBoards
{
    [Authorize(Roles = "Diretoria,Admin,Diretoria")]
    public class DashBoardDiretoriaModel : PageModel
    {
        public List<GestorComercialDash> EquipeIntermedic { get; set; } = new List<GestorComercialDash>();
        public List<GestorComercialDash> EquipeDenuo { get; set; } = new List<GestorComercialDash>();

        public TOTVSDENUOContext ProtheusDenuo { get; set; }
        public TOTVSINTERContext ProtheusInter { get; set; }

        public string Mes { get; set; }
        public string Ano { get; set; }
        public ApplicationDbContext SGID { get; set; }

        public DashBoardDiretoriaModel(TOTVSDENUOContext denuo, TOTVSINTERContext inter, ApplicationDbContext sgid)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
            SGID = sgid;
        }

        public void OnGet()
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                string data = DateTime.Now.ToString("yyyy/MM").Replace("/", "");
                string DataInicio = data + "01";
                string DataFim = data + "31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109 };

                #region Intermedic
                var vendedores = ProtheusInter.Sa3010s.Where(x => x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                {
                    x.A3Nome,
                    x.A3Xlogin,
                }).ToList();

                vendedores.ForEach(x =>
                {
                    EquipeIntermedic.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                });

                #region Valorizadas
                var resultadoValorizado = (from SC50 in ProtheusInter.Sc5010s
                                           join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                           join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                           join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                           join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                           join PA10 in ProtheusInter.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo }
                                           join SUA10 in ProtheusInter.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum }
                                           where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*" && PA10.DELET != "*"
                                           && SUA10.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                           && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                           orderby SC50.C5XDtcir descending
                                           select new RelatorioCirurgiasValorizadas
                                           {
                                               A3XLogin = SA30.A3Xlogin,
                                               C5Num = SC50.C5Num,
                                               C6Valor = SC60.C6Valor,
                                           }
                                       ).GroupBy(x => new
                                       {
                                           x.A3XLogin,
                                           x.C5Num,
                                       }).Select(x => new RelatorioCirurgiasValorizadas
                                       {
                                           A3XLogin = x.Key.A3XLogin,
                                           C5Num = x.Key.C5Num,
                                           C6Valor = x.Sum(c => c.C6Valor)
                                       }).ToList();

                EquipeIntermedic.ForEach(x =>
                {
                    x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                    x.CirurgiasValorizadasValor = resultadoValorizado.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                });

                #endregion

                #region Faturados
                var query = (from SD20 in ProtheusInter.Sd2010s
                             join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                             && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                             (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                             && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                             select new
                             {
                                 Login = SA30.A3Xlogin,
                                 NF = SD20.D2Doc,
                                 Total = SD20.D2Total,

                             });

                var resultado = query.GroupBy(x => new
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

                EquipeIntermedic.ForEach(x =>
                {
                    x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User).DistinctBy(x => x.Nf).Count();
                    x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User).Sum(x => x.Total);

                    var integrante = x.User.ToLower();

                    var time = SGID.Times.FirstOrDefault(x => x.Integrante == integrante);

                    if (time == null)
                    {
                        x.Meta = 0;
                    }
                    else
                    {
                        x.Meta = time.Meta - x.FaturadoMesValor;
                    }

                    if (x.Meta < 0)
                    {
                        x.Meta = 0.0;
                    }
                });
                #endregion

                #region EmAberto
                var resultadoEmAberto = (from SC5 in ProtheusInter.Sc5010s
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
                                         select new RelatorioCirurgiasFaturar
                                         {
                                             Login = SA3.A3Xlogin,
                                             Num = SC5.C5Num,
                                             Valor = SC6.C6Valor,
                                         }
                                     ).GroupBy(x => new
                                     {
                                         x.Login,
                                         x.Num,
                                     }).Select(x => new RelatorioCirurgiasFaturar
                                     {
                                         Login = x.Key.Login,
                                         Num = x.Key.Num,
                                         Valor = x.Sum(c => c.Valor)
                                     }).ToList();

                EquipeIntermedic.ForEach(x =>
                {
                    x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                    x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                });
                #endregion
                
                EquipeIntermedic = EquipeIntermedic.OrderBy(x => x.Nome).ToList();

                #endregion

                #region Denuo
                var vendedoresDenuo = ProtheusDenuo.Sa3010s.Where(x => x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                {
                    x.A3Nome,
                    x.A3Xlogin,
                }).ToList();

                vendedoresDenuo.ForEach(x =>
                {
                    EquipeDenuo.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                });

                #region Valorizadas

                var resultadoValorizadoDenuo = (from SC50 in ProtheusDenuo.Sc5010s
                                                join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                join PA10 in ProtheusDenuo.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                                from c in Sr.DefaultIfEmpty()
                                                join SUA10 in ProtheusDenuo.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                                from a in Rs.DefaultIfEmpty()
                                                where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                                && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                                orderby SC50.C5XDtcir descending
                                                select new RelatorioCirurgiasValorizadas
                                                {
                                                    A3XLogin = SA30.A3Xlogin,
                                                    C5Num = SC50.C5Num,
                                                    C6Valor = SC60.C6Valor,
                                                }
                                       ).GroupBy(x => new
                                       {
                                           x.A3XLogin,
                                           x.C5Num,
                                       }).Select(x => new RelatorioCirurgiasValorizadas
                                       {
                                           A3XLogin = x.Key.A3XLogin,
                                           C5Num = x.Key.C5Num,
                                           C6Valor = x.Sum(c => c.C6Valor)
                                       }).ToList();

                EquipeDenuo.ForEach(x =>
                        {
                            x.CirurgiasValorizadas = resultadoValorizadoDenuo.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                            x.CirurgiasValorizadasValor = resultadoValorizadoDenuo.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                        });

                #endregion

                #region EmAberto
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
                                              orderby SC5.C5Num, SC5.C5Emissao descending
                                              select new RelatorioCirurgiasFaturar
                                              {
                                                  Login = SA3.A3Xlogin,
                                                  Num = SC5.C5Num,
                                                  Valor = SC6.C6Valor,
                                              }
                                     ).GroupBy(x => new
                                     {
                                         x.Login,
                                         x.Num,
                                     }).Select(x => new RelatorioCirurgiasFaturar
                                     {
                                         Login = x.Key.Login,
                                         Num = x.Key.Num,
                                         Valor = x.Sum(c => c.Valor)
                                     }).ToList();

                EquipeDenuo.ForEach(x =>
                        {
                            x.CirurgiasEmAberto = resultadoEmAbertoDenuo.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                            x.CirurgiasEmAbertoValor = resultadoEmAbertoDenuo.Where(c => c.Login == x.User).Sum(x => x.Valor);
                        });

                #endregion

                #region Faturado
                var queryDenuo = (from SD20 in ProtheusDenuo.Sd2010s
                                  join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                  join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                  join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                  where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                  && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                  (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
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
                    A3Nome = x.Key.Login,
                    Nf = x.Key.NF,
                    Total = x.Sum(c => c.Total),
                }).ToList();

                EquipeDenuo.ForEach(x =>
                        {
                            x.FaturadoMes = resultadoDenuo.Where(c => c.A3Nome == x.User).DistinctBy(x => x.Nf).Count();
                            x.FaturadoMesValor = resultadoDenuo.Where(c => c.A3Nome == x.User).Sum(x => x.Total);

                            var integrante = x.User.ToLower();

                            var time = SGID.Times.FirstOrDefault(x => x.Integrante == integrante);

                            if (time == null)
                            {
                                x.Meta = 0;
                            }
                            else
                            {
                                x.Meta = time.Meta - x.FaturadoMesValor;
                            }

                            if (x.Meta < 0)
                            {
                                x.Meta = 0.0;
                            }
                        });
                #endregion

                EquipeDenuo = EquipeDenuo.OrderBy(x => x.Nome).ToList();
                #endregion

            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex, SGID, "DashBoardDiretoria", user);
            }
        }

        public IActionResult OnPost(string Ano, string Mes)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                string DataInicio = $"{Ano}{Mes}01";
                string DataFim = $"{Ano}{Mes}31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109 };

                #region Intermedic
                var vendedores = ProtheusInter.Sa3010s.Where(x => x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                {
                    x.A3Nome,
                    x.A3Xlogin,
                }).ToList();

                vendedores.ForEach(x =>
                {
                    EquipeIntermedic.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                });

                var resultadoValorizado = (from SC50 in ProtheusInter.Sc5010s
                                           join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                           join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                           join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                           join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                           join PA10 in ProtheusInter.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo }
                                           join SUA10 in ProtheusInter.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum }
                                           where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*" && PA10.DELET != "*"
                                           && SUA10.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                           && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                           orderby SC50.C5XDtcir descending
                                           select new RelatorioCirurgiasValorizadas
                                           {
                                               A3XLogin = SA30.A3Xlogin,
                                               C5Num = SC50.C5Num,
                                               C6Valor = SC60.C6Valor,
                                           }
                                       ).GroupBy(x => new
                                       {
                                           x.A3XLogin,
                                           x.C5Num,
                                       }).Select(x => new RelatorioCirurgiasValorizadas
                                       {
                                           A3XLogin = x.Key.A3XLogin,
                                           C5Num = x.Key.C5Num,
                                           C6Valor = x.Sum(c => c.C6Valor)
                                       }).ToList();

                EquipeIntermedic.ForEach(x =>
                {
                    x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                    x.CirurgiasValorizadasValor = resultadoValorizado.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                });




                var query = (from SD20 in ProtheusInter.Sd2010s
                             join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                             && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                             (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                             && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                             select new
                             {
                                 Login = SA30.A3Xlogin,
                                 NF = SD20.D2Doc,
                                 Total = SD20.D2Total,

                             });

                var resultado = query.GroupBy(x => new
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

                EquipeIntermedic.ForEach(x =>
                {
                    x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User).DistinctBy(x => x.Nf).Count();
                    x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User).Sum(x => x.Total);

                    var integrante = x.User.ToLower();

                    var time = SGID.Times.FirstOrDefault(x => x.Integrante == integrante);

                    if (time == null)
                    {
                        x.Meta = 0;
                    }
                    else
                    {
                        x.Meta = time.Meta - x.FaturadoMesValor;
                    }

                    if (x.Meta < 0)
                    {
                        x.Meta = 0.0;
                    }
                });


                var resultadoEmAberto = (from SC5 in ProtheusInter.Sc5010s
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
                                         select new RelatorioCirurgiasFaturar
                                         {
                                             Login = SA3.A3Xlogin,
                                             Num = SC5.C5Num,
                                             Valor = SC6.C6Valor,
                                         }
                                     ).GroupBy(x => new
                                     {
                                         x.Login,
                                         x.Num,
                                     }).Select(x => new RelatorioCirurgiasFaturar
                                     {
                                         Login = x.Key.Login,
                                         Num = x.Key.Num,
                                         Valor = x.Sum(c => c.Valor)
                                     }).ToList();

                EquipeIntermedic.ForEach(x =>
                {
                    x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                    x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                });

                EquipeIntermedic = EquipeIntermedic.OrderBy(x => x.Nome).ToList();

                #endregion

                #region Denuo
                var vendedoresDenuo = ProtheusDenuo.Sa3010s.Where(x => x.A3Xlogsup != "" && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                {
                    x.A3Nome,
                    x.A3Xlogin,
                }).ToList();

                vendedoresDenuo.ForEach(x =>
                {
                    EquipeDenuo.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                });

                var resultadoValorizadoDenuo = (from SC50 in ProtheusDenuo.Sc5010s
                                                join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                join PA10 in ProtheusDenuo.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                                from c in Sr.DefaultIfEmpty()
                                                join SUA10 in ProtheusDenuo.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                                from a in Rs.DefaultIfEmpty()
                                                where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                                && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                                orderby SC50.C5XDtcir descending
                                                select new RelatorioCirurgiasValorizadas
                                                {
                                                    A3XLogin = SA30.A3Xlogin,
                                                    C5Num = SC50.C5Num,
                                                    C6Valor = SC60.C6Valor,
                                                }
                                       ).GroupBy(x => new
                                       {
                                           x.A3XLogin,
                                           x.C5Num,
                                       }).Select(x => new RelatorioCirurgiasValorizadas
                                       {
                                           A3XLogin = x.Key.A3XLogin,
                                           C5Num = x.Key.C5Num,
                                           C6Valor = x.Sum(c => c.C6Valor)
                                       }).ToList();

                EquipeDenuo.ForEach(x =>
                {
                    x.CirurgiasValorizadas = resultadoValorizadoDenuo.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                    x.CirurgiasValorizadasValor = resultadoValorizadoDenuo.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                });

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
                                              && SA3.A3Xlogsup != "ANDRE.SALES"
                                              orderby SC5.C5Num, SC5.C5Emissao descending
                                              select new RelatorioCirurgiasFaturar
                                              {
                                                  Login = SA3.A3Xlogin,
                                                  Num = SC5.C5Num,
                                                  Valor = SC6.C6Valor,
                                              }
                                     ).GroupBy(x => new
                                     {
                                         x.Login,
                                         x.Num,
                                     }).Select(x => new RelatorioCirurgiasFaturar
                                     {
                                         Login = x.Key.Login,
                                         Num = x.Key.Num,
                                         Valor = x.Sum(c => c.Valor)
                                     }).ToList();

                EquipeDenuo.ForEach(x =>
                {
                    x.CirurgiasEmAberto = resultadoEmAbertoDenuo.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                    x.CirurgiasEmAbertoValor = resultadoEmAbertoDenuo.Where(c => c.Login == x.User).Sum(x => x.Valor);
                });

                var queryDenuo = (from SD20 in ProtheusDenuo.Sd2010s
                                  join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                  join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                  join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                  where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                  && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                  (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
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
                    A3Nome = x.Key.Login,
                    Nf = x.Key.NF,
                    Total = x.Sum(c => c.Total),
                }).ToList();

                EquipeDenuo.ForEach(x =>
                {
                    x.FaturadoMes = resultadoDenuo.Where(c => c.A3Nome == x.User).DistinctBy(x => x.Nf).Count();
                    x.FaturadoMesValor = resultadoDenuo.Where(c => c.A3Nome == x.User).Sum(x => x.Total);

                    var integrante = x.User.ToLower();

                    var time = SGID.Times.FirstOrDefault(x => x.Integrante == integrante);

                    if (time == null)
                    {
                        x.Meta = 0;
                    }
                    else
                    {
                        x.Meta = time.Meta - x.FaturadoMesValor;
                    }

                    if (x.Meta < 0)
                    {
                        x.Meta = 0.0;
                    }
                });
                EquipeDenuo = EquipeDenuo.OrderBy(x => x.Nome).ToList();
                #endregion

            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex, SGID, "DashBoardDiretoria", user);

            }
            return Page();
        }
    }
}
