using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Migrations;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Comercial;
using SGID.Models.Denuo;
using System.Linq;
using SGID.Models.Financeiro;

namespace SGID.Pages.DashBoards
{
    [Authorize]
    public class DashBoardGestorComercialModel : PageModel
    {
        public List<GestorComercialDash> Equipe { get; set; } = new List<GestorComercialDash>();

        public TOTVSDENUOContext ProtheusDenuo { get; set; }
        public TOTVSINTERContext ProtheusInter { get; set; }

        public string Mes { get; set; }
        public string Ano { get; set; }
        public ApplicationDbContext SGID { get; set; }

        public DashBoardGestorComercialModel(TOTVSDENUOContext denuo, TOTVSINTERContext inter, ApplicationDbContext sgid)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
            SGID = sgid;
        }

        public void OnGet(string Id)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                string data = DateTime.Now.ToString("yyyy/MM").Replace("/", "");
                string DataInicio = data + "01";
                string DataFim = data + "31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109 };

                if (Id == "01")
                {
                    if (Convert.ToInt32(DataInicio) < 20240228)
                    {
                        var vendedores = ProtheusInter.Sa3010s.Where(x => (x.A3Xlogsup == user || x.A3Xlogin == user) && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                        {
                            x.A3Nome,
                            x.A3Xlogin,
                        }).ToList();

                        vendedores.ForEach(x =>
                        {
                            Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                        });

                        #region Valorizado

                        var resultadoValorizado = (from SC50 in ProtheusInter.Sc5010s
                                                   join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                   join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                   join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                   join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                   join PA10 in ProtheusInter.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo }
                                                   join SUA10 in ProtheusInter.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum }
                                                   where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && PA10.DELET != "*"
                                                   && SUA10.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                   && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user) && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                        Equipe.ForEach(x =>
                        {
                            x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                            x.CirurgiasValorizadasValor = resultadoValorizado.Sum(x => x.C6Valor);
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
                                                 && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140" &&
                                                 (SA3.A3Xlogsup == user || SA3.A3Xlogin == user)
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

                        Equipe.ForEach(x =>
                        {
                            x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                            x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                        });
                        #endregion

                        #region Faturamento
                        var query = (from SD20 in ProtheusInter.Sd2010s
                                     join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                     join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                     join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                     where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                     && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                     (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                     && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                     && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                        }).Select(x => new RelatorioCirurgiasFaturadas
                        {
                            A3Nome = x.Key.Login,
                            Nf = x.Key.NF,
                            Total = x.Sum(c => c.Total),
                        }).ToList();
                        #endregion

                        #region Devolução
                        var teste2 = (from SD10 in ProtheusInter.Sd1010s
                                      join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                      join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                      join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                      join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                      where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                      && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                      && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                      && (int)(object)SD10.D1Dtdigit >= 20200701 && SB10.B1Ugrpint != "082"
                                      && SC50.C5Utpoper == "F" && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                                          SA30.A3Xdescun
                                      }
                             ).GroupBy(x => new
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
                                 x.A3Xdescun
                             });

                        var Relatorio = teste2.Select(x => new RelatorioDevolucaoFat
                        {
                            Filial = x.Key.D1Filial,
                            Clifor = x.Key.D1Fornece,
                            Loja = x.Key.D1Loja,
                            Nome = x.Key.A1Nome,
                            Tipo = x.Key.A1Clinter,
                            Nf = x.Key.D1Doc,
                            Serie = x.Key.D1Serie,
                            Digitacao = x.Key.D1Dtdigit,
                            Total = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc),
                            Valipi = x.Sum(c => c.D1Valipi),
                            Valicm = x.Sum(c => c.D1Valicm),
                            Descon = x.Sum(c => c.D1Valdesc),
                            TotalBrut = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valdesc),
                            A3Nome = x.Key.A3Nome,
                            D1Nfori = x.Key.D1Nfori,
                            D1Seriori = x.Key.D1Seriori,
                            D1Datori = x.Key.D1Datori,
                            Linha = x.Key.A3Xdescun
                        }).ToList();
                        #endregion

                        Equipe.ForEach(x =>
                        {
                            x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User).DistinctBy(x => x.Nf).Count();
                            x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User).Sum(x => x.Total) - Relatorio.Where(c => c.A3Nome == x.User).Sum(x => x.Total);

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


                    }
                    else
                    {
                        var vendedores = ProtheusInter.Sa3010s.Where(x => (x.A3Xlogsup == user || x.A3Xlogin == user) && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                        {
                            x.A3Nome,
                            x.A3Xlogin,
                        }).ToList();

                        vendedores.ForEach(x =>
                        {
                            Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                        });

                        #region Valorizado

                        var resultadoValorizado = (from SC50 in ProtheusInter.Sc5010s
                                                   join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                   join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                   join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                   join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                   join PA10 in ProtheusInter.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo }
                                                   join SUA10 in ProtheusInter.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum }
                                                   where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && PA10.DELET != "*"
                                                   && SUA10.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                   && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user) && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                        Equipe.ForEach(x =>
                        {
                            x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                            x.CirurgiasValorizadasValor = resultadoValorizado.Sum(x => x.C6Valor);
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
                                                 && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140" &&
                                                 (SA3.A3Xlogsup == user || SA3.A3Xlogin == user)
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

                        Equipe.ForEach(x =>
                        {
                            x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                            x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                        });
                        #endregion

                        #region BaixaLicitacoes

                        var BaixaLicitacoesInter = (from SE50 in ProtheusInter.Se5010s
                                                    join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                                    equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                                    join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                                    join SC50 in ProtheusInter.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                                    join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                    join SD20 in ProtheusInter.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                                    where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                                    && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                                    && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                                    && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                                    && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                                    && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                                    && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
                                                    && (int)(object)SD20.D2Emissao >= 20240301
                                                    && SA30.A3Xlogsup != "ANDRE.SALES"
                                                    select new RelatorioAreceberBaixa
                                                    {
                                                        Prefixo = SE50.E5Prefixo,
                                                        Numero = SE50.E5Numero,
                                                        Parcela = SE50.E5Parcela,
                                                        TP = SE50.E5Tipo,
                                                        CliFor = SE50.E5Clifor,
                                                        NomeFor = SA10.A1Nome,
                                                        Naturez = SE50.E5Naturez,
                                                        Vencimento = SE10.E1Vencto,
                                                        Historico = SE50.E5Histor,
                                                        DataBaixa = SE50.E5Data,
                                                        ValorOrig = SE10.E1Valor,
                                                        JurMulta = SE50.E5Vljuros + SE50.E5Vlmulta,
                                                        Correcao = SE50.E5Vlcorre,
                                                        Descon = SE50.E5Vldesco,
                                                        Abatimento = 0,
                                                        Imposto = 0,
                                                        ValorAcess = 0,
                                                        TotalBaixado = SE50.E5Valor,
                                                        Banco = SE50.E5Banco,
                                                        DtDigi = SE50.E5Dtdigit,
                                                        Mot = SE50.E5Motbx,
                                                        Orig = SE50.E5Filorig,
                                                        Vendedor = SC50.C5Nomvend,
                                                        TipoCliente = SA10.A1Clinter,
                                                        CodigoCliente = SA10.A1Xgrinte,
                                                        Login = SA30.A3Xlogin,
                                                        Gestor = SA30.A3Xlogsup,
                                                        DataPedido = SD20.D2Emissao
                                                    }).GroupBy(x => new
                                                    {
                                                        x.Prefixo,
                                                        x.Numero,
                                                        x.Parcela,
                                                        x.TP,
                                                        x.CliFor,
                                                        x.NomeFor,
                                                        x.Naturez,
                                                        x.Vencimento,
                                                        x.Historico,
                                                        x.DataBaixa,
                                                        x.ValorOrig,
                                                        x.JurMulta,
                                                        x.Correcao,
                                                        x.Descon,
                                                        x.Abatimento,
                                                        x.Imposto,
                                                        x.ValorAcess,
                                                        x.TotalBaixado,
                                                        x.Banco,
                                                        x.DtDigi,
                                                        x.Mot,
                                                        x.Orig,
                                                        x.Vendedor,
                                                        x.TipoCliente,
                                                        x.CodigoCliente,
                                                        x.Login,
                                                        x.Gestor,
                                                        x.DataPedido
                                                    }).Select(x => new RelatorioAreceberBaixa
                                                    {
                                                        Prefixo = x.Key.Prefixo,
                                                        Numero = x.Key.Numero,
                                                        Parcela = x.Key.Parcela,
                                                        TP = x.Key.TP,
                                                        CliFor = x.Key.CliFor,
                                                        NomeFor = x.Key.NomeFor,
                                                        Naturez = x.Key.Naturez,
                                                        Vencimento = x.Key.Vencimento,
                                                        Historico = x.Key.Historico,
                                                        DataBaixa = x.Key.DataBaixa,
                                                        ValorOrig = x.Key.ValorOrig,
                                                        JurMulta = x.Key.JurMulta,
                                                        Correcao = x.Key.Correcao,
                                                        Descon = x.Key.Descon,
                                                        Abatimento = x.Key.Abatimento,
                                                        Imposto = x.Key.Imposto,
                                                        ValorAcess = x.Key.ValorAcess,
                                                        TotalBaixado = x.Key.TotalBaixado,
                                                        Banco = x.Key.Banco,
                                                        DtDigi = x.Key.DtDigi,
                                                        Mot = x.Key.Mot,
                                                        Orig = x.Key.Orig,
                                                        Vendedor = x.Key.Vendedor,
                                                        TipoCliente = x.Key.TipoCliente,
                                                        CodigoCliente = x.Key.CodigoCliente,
                                                        Login = x.Key.Login,
                                                        Gestor = x.Key.Gestor,
                                                        DataPedido = x.Key.DataPedido
                                                    }).ToList();

                        #endregion

                        #region Faturamento
                        var query = (from SD20 in ProtheusInter.Sd2010s
                                     join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                     join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                     join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                     where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                     && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                     (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                     && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                     && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
                                     select new
                                     {
                                         Login = SA30.A3Xlogin,
                                         NF = SD20.D2Doc,
                                         Total = SD20.D2Total,
                                         SA10.A1Xgrinte

                                     });

                        var resultado = query.GroupBy(x => new
                        {
                            x.Login,
                            x.NF,
                            x.A1Xgrinte
                        }).Select(x => new RelatorioCirurgiasFaturadas
                        {
                            A3Nome = x.Key.Login,
                            Nf = x.Key.NF,
                            Total = x.Sum(c => c.Total),
                            Linha = x.Key.A1Xgrinte
                        }).ToList();
                        #endregion

                        #region Devolução
                        var teste2 = (from SD10 in ProtheusInter.Sd1010s
                                      join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                      join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                      join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                      join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                      where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                      && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                      && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                      && (int)(object)SD10.D1Dtdigit >= 20200701 && SB10.B1Ugrpint != "082"
                                      && SC50.C5Utpoper == "F" && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                                          SA30.A3Xdescun
                                      }
                             ).GroupBy(x => new
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
                                 x.A3Xdescun
                             });

                        var Relatorio = teste2.Select(x => new RelatorioDevolucaoFat
                        {
                            Filial = x.Key.D1Filial,
                            Clifor = x.Key.D1Fornece,
                            Loja = x.Key.D1Loja,
                            Nome = x.Key.A1Nome,
                            Tipo = x.Key.A1Clinter,
                            Nf = x.Key.D1Doc,
                            Serie = x.Key.D1Serie,
                            Digitacao = x.Key.D1Dtdigit,
                            Total = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc),
                            Valipi = x.Sum(c => c.D1Valipi),
                            Valicm = x.Sum(c => c.D1Valicm),
                            Descon = x.Sum(c => c.D1Valdesc),
                            TotalBrut = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valdesc),
                            A3Nome = x.Key.A3Nome,
                            D1Nfori = x.Key.D1Nfori,
                            D1Seriori = x.Key.D1Seriori,
                            D1Datori = x.Key.D1Datori,
                            Linha = x.Key.A3Xdescun
                        }).ToList();
                        #endregion

                        Equipe.ForEach(x =>
                        {
                            x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").DistinctBy(x => x.Nf).Count();
                            x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total) - Relatorio.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total);

                            x.CirurgiasLicitacoes = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).DistinctBy(x => x.Nf).Count();
                            x.CirurgiasLicitacoesValor = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total) - Relatorio.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total);

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

                    }

                }
                else
                {

                    if (user != "TIAGO.FONSECA")
                    {
                        if (Convert.ToInt32(DataInicio) < 20240228)
                        {
                            #region Gestor

                            if (user == "ARTEMIO.COSTA") user = "LEONARDO.BRITO";

                            var vendedores = ProtheusDenuo.Sa3010s.Where(x => (x.A3Xlogsup == user || x.A3Xlogin == user) && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                            {
                                x.A3Nome,
                                x.A3Xlogin,
                            }).ToList();

                            vendedores.ForEach(x =>
                            {
                                Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                            });

                            #region Valorizado
                            var resultadoValorizado = (from SC50 in ProtheusDenuo.Sc5010s
                                                       join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                       join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                       join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                       join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                       join PA10 in ProtheusDenuo.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                                       from c in Sr.DefaultIfEmpty()
                                                       join SUA10 in ProtheusDenuo.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                                       from a in Rs.DefaultIfEmpty()
                                                       where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                                       && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                       && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user) && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                                x.CirurgiasValorizadasValor = resultadoValorizado.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                            });
                            #endregion

                            #region EmAberto
                            var resultadoEmAberto = (from SC5 in ProtheusDenuo.Sc5010s
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
                                                     && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140" &&
                                                     (SA3.A3Xlogsup == user || SA3.A3Xlogin == user)
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                                x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                            });
                            #endregion

                            #region Faturado

                            var query = (from SD20 in ProtheusDenuo.Sd2010s
                                         join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                         join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                         join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                         && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                         && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                            }).Select(x => new RelatorioCirurgiasFaturadas
                            {
                                A3Nome = x.Key.Login,
                                Nf = x.Key.NF,
                                Total = x.Sum(c => c.Total),
                            }).ToList();
                            #endregion

                            #region Devolucao
                            var teste = (from SD10 in ProtheusInter.Sd1010s
                                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                         join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                         join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                         where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                         && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                         && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                         && (int)(object)SD10.D1Dtdigit >= 20200701 && SC50.C5Utpoper == "F" && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                                             SA10.A1Xgrinte
                                         }
                                         ).ToList();

                            var RelatorioDev = new List<RelatorioDevolucaoFat>();

                            if (teste.Count != 0)
                            {
                                teste.ForEach(x =>
                                {
                                    if (!RelatorioDev.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                                    {

                                        var Iguais = teste
                                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                                        double desconto = 0;
                                        double valipi = 0;
                                        double total = 0;
                                        double valicm = 0;
                                        Iguais.ForEach(x =>
                                        {
                                            desconto += x.D1Valdesc;
                                            valipi += x.D1Valipi;
                                            total += x.D1Total;
                                            valicm += x.D1Valicm;
                                        });

                                        RelatorioDev.Add(new RelatorioDevolucaoFat
                                        {
                                            Filial = x.D1Filial,
                                            Clifor = x.D1Fornece,
                                            Loja = x.D1Loja,
                                            Nome = x.A1Nome,
                                            Tipo = x.A1Clinter,
                                            Nf = x.D1Doc,
                                            Serie = x.D1Serie,
                                            Digitacao = x.D1Dtdigit,
                                            Total = total - desconto,
                                            Valipi = valipi,
                                            Valicm = valicm,
                                            Descon = desconto,
                                            TotalBrut = total - desconto + valipi,
                                            A3Nome = x.A3Nome,
                                            D1Nfori = x.D1Nfori,
                                            D1Seriori = x.D1Seriori,
                                            D1Datori = x.D1Datori,
                                            Linha = x.A1Xgrinte
                                        });
                                    }
                                });
                            }
                            #endregion

                            Equipe.ForEach(x =>
                            {
                                x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User).DistinctBy(x => x.Nf).Count();
                                x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User).Sum(x => x.Total) - RelatorioDev.Where(c => c.Login == x.User).Sum(x => x.Total);

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
                        }
                        else
                        {
                            #region Gestor

                            if (user == "ARTEMIO.COSTA") user = "LEONARDO.BRITO";

                            var vendedores = ProtheusDenuo.Sa3010s.Where(x => (x.A3Xlogsup == user || x.A3Xlogin == user) && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                            {
                                x.A3Nome,
                                x.A3Xlogin,
                            }).ToList();

                            vendedores.ForEach(x =>
                            {
                                Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                            });

                            #region Valorizados
                            var resultadoValorizado = (from SC50 in ProtheusDenuo.Sc5010s
                                                       join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                       join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                       join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                       join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                       join PA10 in ProtheusDenuo.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                                       from c in Sr.DefaultIfEmpty()
                                                       join SUA10 in ProtheusDenuo.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                                       from a in Rs.DefaultIfEmpty()
                                                       where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                                       && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                       && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user) && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                                x.CirurgiasValorizadasValor = resultadoValorizado.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                            });
                            #endregion

                            #region EmAberto

                            var resultadoEmAberto = (from SC5 in ProtheusDenuo.Sc5010s
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
                                                     && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140" &&
                                                     (SA3.A3Xlogsup == user || SA3.A3Xlogin == user)
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                                x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                            });

                            #endregion

                            #region Faturamento

                            var query = (from SD20 in ProtheusDenuo.Sd2010s
                                         join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                         join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                         join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                         && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                         && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
                                         select new
                                         {
                                             Login = SA30.A3Xlogin,
                                             NF = SD20.D2Doc,
                                             Total = SD20.D2Total,
                                             SA10.A1Xgrinte
                                         });

                            var resultado = query.GroupBy(x => new
                            {
                                x.Login,
                                x.NF,
                                x.A1Xgrinte
                            }).Select(x => new RelatorioCirurgiasFaturadas
                            {
                                A3Nome = x.Key.Login,
                                Nf = x.Key.NF,
                                Total = x.Sum(c => c.Total),
                                Linha = x.Key.A1Xgrinte
                            }).ToList();

                            #endregion

                            #region Devolucao
                            var teste = (from SD10 in ProtheusInter.Sd1010s
                                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                         join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                         join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                         where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                         && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                         && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                         && (int)(object)SD10.D1Dtdigit >= 20200701 && SC50.C5Utpoper == "F" && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                                             SA10.A1Xgrinte
                                         }
                                         ).ToList();

                            var RelatorioDev = new List<RelatorioDevolucaoFat>();

                            if (teste.Count != 0)
                            {
                                teste.ForEach(x =>
                                {
                                    if (!RelatorioDev.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                                    {

                                        var Iguais = teste
                                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                                        double desconto = 0;
                                        double valipi = 0;
                                        double total = 0;
                                        double valicm = 0;
                                        Iguais.ForEach(x =>
                                        {
                                            desconto += x.D1Valdesc;
                                            valipi += x.D1Valipi;
                                            total += x.D1Total;
                                            valicm += x.D1Valicm;
                                        });

                                        RelatorioDev.Add(new RelatorioDevolucaoFat
                                        {
                                            Filial = x.D1Filial,
                                            Clifor = x.D1Fornece,
                                            Loja = x.D1Loja,
                                            Nome = x.A1Nome,
                                            Tipo = x.A1Clinter,
                                            Nf = x.D1Doc,
                                            Serie = x.D1Serie,
                                            Digitacao = x.D1Dtdigit,
                                            Total = total - desconto,
                                            Valipi = valipi,
                                            Valicm = valicm,
                                            Descon = desconto,
                                            TotalBrut = total - desconto + valipi,
                                            A3Nome = x.A3Nome,
                                            D1Nfori = x.D1Nfori,
                                            D1Seriori = x.D1Seriori,
                                            D1Datori = x.D1Datori,
                                            Linha = x.A1Xgrinte
                                        });
                                    }
                                });
                            }
                            #endregion

                            #region BaixaLicitacoes

                            var BaixaLicitacoesDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                                        join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                                        equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                                        join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                                        join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                                        join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                        join SD20 in ProtheusDenuo.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                                        where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                                        && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                                        && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                                        && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                                        && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                                        && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                                        && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
                                                        && (int)(object)SD20.D2Emissao >= 20240301
                                                        && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
                                                        select new RelatorioAreceberBaixa
                                                        {
                                                            Prefixo = SE50.E5Prefixo,
                                                            Numero = SE50.E5Numero,
                                                            Parcela = SE50.E5Parcela,
                                                            TP = SE50.E5Tipo,
                                                            CliFor = SE50.E5Clifor,
                                                            NomeFor = SA10.A1Nome,
                                                            Naturez = SE50.E5Naturez,
                                                            Vencimento = SE10.E1Vencto,
                                                            Historico = SE50.E5Histor,
                                                            DataBaixa = SE50.E5Data,
                                                            ValorOrig = SE10.E1Valor,
                                                            JurMulta = SE50.E5Vljuros + SE50.E5Vlmulta,
                                                            Correcao = SE50.E5Vlcorre,
                                                            Descon = SE50.E5Vldesco,
                                                            Abatimento = 0,
                                                            Imposto = 0,
                                                            ValorAcess = 0,
                                                            TotalBaixado = SE50.E5Valor,
                                                            Banco = SE50.E5Banco,
                                                            DtDigi = SE50.E5Dtdigit,
                                                            Mot = SE50.E5Motbx,
                                                            Orig = SE50.E5Filorig,
                                                            Vendedor = SC50.C5Nomvend,
                                                            TipoCliente = SA10.A1Clinter,
                                                            CodigoCliente = SA10.A1Xgrinte,
                                                            Login = SA30.A3Xlogin,
                                                            Gestor = SA30.A3Xlogsup,
                                                            DataPedido = SD20.D2Emissao
                                                        }).GroupBy(x => new
                                                        {
                                                            x.Prefixo,
                                                            x.Numero,
                                                            x.Parcela,
                                                            x.TP,
                                                            x.CliFor,
                                                            x.NomeFor,
                                                            x.Naturez,
                                                            x.Vencimento,
                                                            x.Historico,
                                                            x.DataBaixa,
                                                            x.ValorOrig,
                                                            x.JurMulta,
                                                            x.Correcao,
                                                            x.Descon,
                                                            x.Abatimento,
                                                            x.Imposto,
                                                            x.ValorAcess,
                                                            x.TotalBaixado,
                                                            x.Banco,
                                                            x.DtDigi,
                                                            x.Mot,
                                                            x.Orig,
                                                            x.Vendedor,
                                                            x.TipoCliente,
                                                            x.CodigoCliente,
                                                            x.Login,
                                                            x.Gestor,
                                                            x.DataPedido
                                                        }).Select(x => new RelatorioAreceberBaixa
                                                        {
                                                            Prefixo = x.Key.Prefixo,
                                                            Numero = x.Key.Numero,
                                                            Parcela = x.Key.Parcela,
                                                            TP = x.Key.TP,
                                                            CliFor = x.Key.CliFor,
                                                            NomeFor = x.Key.NomeFor,
                                                            Naturez = x.Key.Naturez,
                                                            Vencimento = x.Key.Vencimento,
                                                            Historico = x.Key.Historico,
                                                            DataBaixa = x.Key.DataBaixa,
                                                            ValorOrig = x.Key.ValorOrig,
                                                            JurMulta = x.Key.JurMulta,
                                                            Correcao = x.Key.Correcao,
                                                            Descon = x.Key.Descon,
                                                            Abatimento = x.Key.Abatimento,
                                                            Imposto = x.Key.Imposto,
                                                            ValorAcess = x.Key.ValorAcess,
                                                            TotalBaixado = x.Key.TotalBaixado,
                                                            Banco = x.Key.Banco,
                                                            DtDigi = x.Key.DtDigi,
                                                            Mot = x.Key.Mot,
                                                            Orig = x.Key.Orig,
                                                            Vendedor = x.Key.Vendedor,
                                                            TipoCliente = x.Key.TipoCliente,
                                                            CodigoCliente = x.Key.CodigoCliente,
                                                            Login = x.Key.Login,
                                                            Gestor = x.Key.Gestor,
                                                            DataPedido = x.Key.DataPedido
                                                        }).ToList();

                            #endregion


                            Equipe.ForEach(x =>
                            {
                                x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").DistinctBy(x => x.Nf).Count();
                                x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total) - RelatorioDev.Where(c => c.Login == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total);

                                x.CirurgiasLicitacoes = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).DistinctBy(x => x.Nf).Count();
                                x.CirurgiasLicitacoesValor = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total) - RelatorioDev.Where(c => c.Login == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total);

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
                        }
                    }
                    else
                    {

                        if (Convert.ToInt32(DataInicio) < 20240228)
                        {
                            #region Tiago

                            var vendedores = ProtheusDenuo.Sa3010s.Where(x => x.A3Xlogsup != "ANDRE.SALES" && x.A3Xlogsup != "" && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                            {
                                x.A3Nome,
                                x.A3Xlogin,
                            }).ToList();

                            vendedores.ForEach(x =>
                            {
                                Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                            });

                            #region Valorizado
                            var resultadoValorizado = (from SC50 in ProtheusDenuo.Sc5010s
                                                       join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                       join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                       join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                       join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                       join PA10 in ProtheusDenuo.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                                       from c in Sr.DefaultIfEmpty()
                                                       join SUA10 in ProtheusDenuo.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                                       from a in Rs.DefaultIfEmpty()
                                                       where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                                       && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                       && SA30.A3Xlogsup != "ANDRE.SALES" && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                                x.CirurgiasValorizadasValor = resultadoValorizado.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                            });
                            #endregion

                            #region EmAberto
                            var resultadoEmAberto = (from SC5 in ProtheusDenuo.Sc5010s
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                                x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                            });
                            #endregion

                            #region Faturamento
                            var query = (from SD20 in ProtheusDenuo.Sd2010s
                                         join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                         join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                         join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                         && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                         && SA30.A3Xlogsup != "ANDRE.SALES"
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
                            }).Select(x => new RelatorioCirurgiasFaturadas
                            {
                                A3Nome = x.Key.Login,
                                Nf = x.Key.NF,
                                Total = x.Sum(c => c.Total),
                            }).ToList();
                            #endregion

                            #region Devolucao
                            var teste = (from SD10 in ProtheusInter.Sd1010s
                                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                         join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                         join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                         where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                         && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                         && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                         && (int)(object)SD10.D1Dtdigit >= 20200701 && SC50.C5Utpoper == "F" && SA30.A3Xlogsup != "ANDRE.SALES"
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
                                             SA10.A1Xgrinte
                                         }
                                         ).ToList();

                            var RelatorioDev = new List<RelatorioDevolucaoFat>();

                            if (teste.Count != 0)
                            {
                                teste.ForEach(x =>
                                {
                                    if (!RelatorioDev.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                                    {

                                        var Iguais = teste
                                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                                        double desconto = 0;
                                        double valipi = 0;
                                        double total = 0;
                                        double valicm = 0;
                                        Iguais.ForEach(x =>
                                        {
                                            desconto += x.D1Valdesc;
                                            valipi += x.D1Valipi;
                                            total += x.D1Total;
                                            valicm += x.D1Valicm;
                                        });

                                        RelatorioDev.Add(new RelatorioDevolucaoFat
                                        {
                                            Filial = x.D1Filial,
                                            Clifor = x.D1Fornece,
                                            Loja = x.D1Loja,
                                            Nome = x.A1Nome,
                                            Tipo = x.A1Clinter,
                                            Nf = x.D1Doc,
                                            Serie = x.D1Serie,
                                            Digitacao = x.D1Dtdigit,
                                            Total = total - desconto,
                                            Valipi = valipi,
                                            Valicm = valicm,
                                            Descon = desconto,
                                            TotalBrut = total - desconto + valipi,
                                            A3Nome = x.A3Nome,
                                            D1Nfori = x.D1Nfori,
                                            D1Seriori = x.D1Seriori,
                                            D1Datori = x.D1Datori,
                                            Linha = x.A1Xgrinte
                                        });
                                    }
                                });
                            }
                            #endregion

                            Equipe.ForEach(x =>
                            {
                                x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User).DistinctBy(x => x.Nf).Count();
                                x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User).Sum(x => x.Total) - RelatorioDev.Where(c => c.A3Nome == x.User).Sum(x => x.Total);

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
                        }
                        else
                        {
                            #region Tiago

                            var vendedores = ProtheusDenuo.Sa3010s.Where(x => x.A3Xlogsup != "ANDRE.SALES" && x.A3Xlogsup != "" && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                            {
                                x.A3Nome,
                                x.A3Xlogin,
                            }).ToList();

                            vendedores.ForEach(x =>
                            {
                                Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                            });

                            #region Valorizado
                            var resultadoValorizado = (from SC50 in ProtheusDenuo.Sc5010s
                                                       join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                       join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                       join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                       join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                       join PA10 in ProtheusDenuo.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                                       from c in Sr.DefaultIfEmpty()
                                                       join SUA10 in ProtheusDenuo.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                                       from a in Rs.DefaultIfEmpty()
                                                       where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                                       && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                       && SA30.A3Xlogsup != "ANDRE.SALES" && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                                x.CirurgiasValorizadasValor = resultadoValorizado.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                            });
                            #endregion

                            #region EmAberto
                            var resultadoEmAberto = (from SC5 in ProtheusDenuo.Sc5010s
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                                x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                            });
                            #endregion

                            #region Faturamento
                            var query = (from SD20 in ProtheusDenuo.Sd2010s
                                         join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                         join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                         join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                         && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                         && SA30.A3Xlogsup != "ANDRE.SALES"
                                         select new
                                         {
                                             Login = SA30.A3Xlogin,
                                             NF = SD20.D2Doc,
                                             Total = SD20.D2Total,
                                             SA10.A1Xgrinte
                                         });

                            var resultado = query.GroupBy(x => new
                            {
                                x.Login,
                                x.NF,
                                x.A1Xgrinte
                            }).Select(x => new RelatorioCirurgiasFaturadas
                            {
                                A3Nome = x.Key.Login,
                                Nf = x.Key.NF,
                                Total = x.Sum(c => c.Total),
                                Linha = x.Key.A1Xgrinte
                            }).ToList();
                            #endregion

                            #region Devolucao
                            var teste = (from SD10 in ProtheusInter.Sd1010s
                                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                         join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                         join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                         where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                         && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                         && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                         && (int)(object)SD10.D1Dtdigit >= 20200701 && SC50.C5Utpoper == "F" && SA30.A3Xlogsup != "ANDRE.SALES"
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
                                             SA10.A1Xgrinte
                                         }
                                         ).ToList();

                            var RelatorioDev = new List<RelatorioDevolucaoFat>();

                            if (teste.Count != 0)
                            {
                                teste.ForEach(x =>
                                {
                                    if (!RelatorioDev.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                                    {

                                        var Iguais = teste
                                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                                        double desconto = 0;
                                        double valipi = 0;
                                        double total = 0;
                                        double valicm = 0;
                                        Iguais.ForEach(x =>
                                        {
                                            desconto += x.D1Valdesc;
                                            valipi += x.D1Valipi;
                                            total += x.D1Total;
                                            valicm += x.D1Valicm;
                                        });

                                        RelatorioDev.Add(new RelatorioDevolucaoFat
                                        {
                                            Filial = x.D1Filial,
                                            Clifor = x.D1Fornece,
                                            Loja = x.D1Loja,
                                            Nome = x.A1Nome,
                                            Tipo = x.A1Clinter,
                                            Nf = x.D1Doc,
                                            Serie = x.D1Serie,
                                            Digitacao = x.D1Dtdigit,
                                            Total = total - desconto,
                                            Valipi = valipi,
                                            Valicm = valicm,
                                            Descon = desconto,
                                            TotalBrut = total - desconto + valipi,
                                            A3Nome = x.A3Nome,
                                            D1Nfori = x.D1Nfori,
                                            D1Seriori = x.D1Seriori,
                                            D1Datori = x.D1Datori,
                                            Linha = x.A1Xgrinte
                                        });
                                    }
                                });
                            }
                            #endregion

                            #region BaixaLicitacoes

                            var BaixaLicitacoesDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                                        join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                                        equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                                        join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                                        join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                                        join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                        join SD20 in ProtheusDenuo.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                                        where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                                        && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                                        && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                                        && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                                        && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                                        && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                                        && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
                                                        && (int)(object)SD20.D2Emissao >= 20240301
                                                        && SA30.A3Xlogsup != "ANDRE.SALES"
                                                        select new RelatorioAreceberBaixa
                                                        {
                                                            Prefixo = SE50.E5Prefixo,
                                                            Numero = SE50.E5Numero,
                                                            Parcela = SE50.E5Parcela,
                                                            TP = SE50.E5Tipo,
                                                            CliFor = SE50.E5Clifor,
                                                            NomeFor = SA10.A1Nome,
                                                            Naturez = SE50.E5Naturez,
                                                            Vencimento = SE10.E1Vencto,
                                                            Historico = SE50.E5Histor,
                                                            DataBaixa = SE50.E5Data,
                                                            ValorOrig = SE10.E1Valor,
                                                            JurMulta = SE50.E5Vljuros + SE50.E5Vlmulta,
                                                            Correcao = SE50.E5Vlcorre,
                                                            Descon = SE50.E5Vldesco,
                                                            Abatimento = 0,
                                                            Imposto = 0,
                                                            ValorAcess = 0,
                                                            TotalBaixado = SE50.E5Valor,
                                                            Banco = SE50.E5Banco,
                                                            DtDigi = SE50.E5Dtdigit,
                                                            Mot = SE50.E5Motbx,
                                                            Orig = SE50.E5Filorig,
                                                            Vendedor = SC50.C5Nomvend,
                                                            TipoCliente = SA10.A1Clinter,
                                                            CodigoCliente = SA10.A1Xgrinte,
                                                            Login = SA30.A3Xlogin,
                                                            Gestor = SA30.A3Xlogsup,
                                                            DataPedido = SD20.D2Emissao
                                                        }).GroupBy(x => new
                                                        {
                                                            x.Prefixo,
                                                            x.Numero,
                                                            x.Parcela,
                                                            x.TP,
                                                            x.CliFor,
                                                            x.NomeFor,
                                                            x.Naturez,
                                                            x.Vencimento,
                                                            x.Historico,
                                                            x.DataBaixa,
                                                            x.ValorOrig,
                                                            x.JurMulta,
                                                            x.Correcao,
                                                            x.Descon,
                                                            x.Abatimento,
                                                            x.Imposto,
                                                            x.ValorAcess,
                                                            x.TotalBaixado,
                                                            x.Banco,
                                                            x.DtDigi,
                                                            x.Mot,
                                                            x.Orig,
                                                            x.Vendedor,
                                                            x.TipoCliente,
                                                            x.CodigoCliente,
                                                            x.Login,
                                                            x.Gestor,
                                                            x.DataPedido
                                                        }).Select(x => new RelatorioAreceberBaixa
                                                        {
                                                            Prefixo = x.Key.Prefixo,
                                                            Numero = x.Key.Numero,
                                                            Parcela = x.Key.Parcela,
                                                            TP = x.Key.TP,
                                                            CliFor = x.Key.CliFor,
                                                            NomeFor = x.Key.NomeFor,
                                                            Naturez = x.Key.Naturez,
                                                            Vencimento = x.Key.Vencimento,
                                                            Historico = x.Key.Historico,
                                                            DataBaixa = x.Key.DataBaixa,
                                                            ValorOrig = x.Key.ValorOrig,
                                                            JurMulta = x.Key.JurMulta,
                                                            Correcao = x.Key.Correcao,
                                                            Descon = x.Key.Descon,
                                                            Abatimento = x.Key.Abatimento,
                                                            Imposto = x.Key.Imposto,
                                                            ValorAcess = x.Key.ValorAcess,
                                                            TotalBaixado = x.Key.TotalBaixado,
                                                            Banco = x.Key.Banco,
                                                            DtDigi = x.Key.DtDigi,
                                                            Mot = x.Key.Mot,
                                                            Orig = x.Key.Orig,
                                                            Vendedor = x.Key.Vendedor,
                                                            TipoCliente = x.Key.TipoCliente,
                                                            CodigoCliente = x.Key.CodigoCliente,
                                                            Login = x.Key.Login,
                                                            Gestor = x.Key.Gestor,
                                                            DataPedido = x.Key.DataPedido
                                                        }).ToList();

                            #endregion

                            Equipe.ForEach(x =>
                            {
                                x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").DistinctBy(x => x.Nf).Count();
                                x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total) - RelatorioDev.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total);

                                x.CirurgiasLicitacoes = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).DistinctBy(x => x.Nf).Count();
                                x.CirurgiasLicitacoesValor = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total) - RelatorioDev.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total);

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
                        }


                    }

                }

                Equipe = Equipe.OrderBy(x => x.Nome).ToList();
            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex, SGID, "DashBoardGestorComercial", user);
            }
        }

        public IActionResult OnPost(string Id,string Ano,string Mes)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                string DataInicio = $"{Ano}{Mes}01";
                string DataFim = $"{Ano}{Mes}31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109 };

                if (Id == "01")
                {
                    if (Convert.ToInt32(DataInicio) < 20240228)
                    {
                        var vendedores = ProtheusInter.Sa3010s.Where(x => (x.A3Xlogsup == user || x.A3Xlogin == user) && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                        {
                            x.A3Nome,
                            x.A3Xlogin,
                        }).ToList();

                        vendedores.ForEach(x =>
                        {
                            Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                        });

                        #region Valorizado

                        var resultadoValorizado = (from SC50 in ProtheusInter.Sc5010s
                                                   join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                   join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                   join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                   join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                   join PA10 in ProtheusInter.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo }
                                                   join SUA10 in ProtheusInter.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum }
                                                   where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && PA10.DELET != "*"
                                                   && SUA10.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                   && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user) && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                        Equipe.ForEach(x =>
                        {
                            x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                            x.CirurgiasValorizadasValor = resultadoValorizado.Sum(x => x.C6Valor);
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
                                                 && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140" &&
                                                 (SA3.A3Xlogsup == user || SA3.A3Xlogin == user)
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

                        Equipe.ForEach(x =>
                        {
                            x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                            x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                        });
                        #endregion

                        #region Faturamento
                        var query = (from SD20 in ProtheusInter.Sd2010s
                                     join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                     join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                     join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                     where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                     && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                     (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                     && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                     && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                        }).Select(x => new RelatorioCirurgiasFaturadas
                        {
                            A3Nome = x.Key.Login,
                            Nf = x.Key.NF,
                            Total = x.Sum(c => c.Total),
                        }).ToList();
                        #endregion

                        #region Devolução
                        var teste2 = (from SD10 in ProtheusInter.Sd1010s
                                     join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                     join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                     join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                     where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                     && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                     && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                     && (int)(object)SD10.D1Dtdigit >= 20200701 && SB10.B1Ugrpint != "082"
                                     && SC50.C5Utpoper == "F" && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                                         SA30.A3Xdescun
                                     }
                             ).GroupBy(x => new
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
                                 x.A3Xdescun
                             });

                        var Relatorio = teste2.Select(x => new RelatorioDevolucaoFat
                        {
                            Filial = x.Key.D1Filial,
                            Clifor = x.Key.D1Fornece,
                            Loja = x.Key.D1Loja,
                            Nome = x.Key.A1Nome,
                            Tipo = x.Key.A1Clinter,
                            Nf = x.Key.D1Doc,
                            Serie = x.Key.D1Serie,
                            Digitacao = x.Key.D1Dtdigit,
                            Total = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc),
                            Valipi = x.Sum(c => c.D1Valipi),
                            Valicm = x.Sum(c => c.D1Valicm),
                            Descon = x.Sum(c => c.D1Valdesc),
                            TotalBrut = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valdesc),
                            A3Nome = x.Key.A3Nome,
                            D1Nfori = x.Key.D1Nfori,
                            D1Seriori = x.Key.D1Seriori,
                            D1Datori = x.Key.D1Datori,
                            Linha = x.Key.A3Xdescun
                        }).ToList();
                        #endregion

                        Equipe.ForEach(x =>
                        {
                            x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User).DistinctBy(x => x.Nf).Count();
                            x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User).Sum(x => x.Total) - Relatorio.Where(c => c.A3Nome == x.User).Sum(x => x.Total);

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


                    }
                    else
                    {
                        var vendedores = ProtheusInter.Sa3010s.Where(x => (x.A3Xlogsup == user || x.A3Xlogin == user) && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                        {
                            x.A3Nome,
                            x.A3Xlogin,
                        }).ToList();

                        vendedores.ForEach(x =>
                        {
                            Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                        });

                        #region Valorizado

                        var resultadoValorizado = (from SC50 in ProtheusInter.Sc5010s
                                                   join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                   join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                   join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                   join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                   join PA10 in ProtheusInter.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo }
                                                   join SUA10 in ProtheusInter.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum }
                                                   where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && PA10.DELET != "*"
                                                   && SUA10.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                   && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user) && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                        Equipe.ForEach(x =>
                        {
                            x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                            x.CirurgiasValorizadasValor = resultadoValorizado.Sum(x => x.C6Valor);
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
                                                 && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140" &&
                                                 (SA3.A3Xlogsup == user || SA3.A3Xlogin == user)
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

                        Equipe.ForEach(x =>
                        {
                            x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                            x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                        });
                        #endregion

                        #region BaixaLicitacoes

                        var BaixaLicitacoesInter = (from SE50 in ProtheusInter.Se5010s
                                                    join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                                    equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                                    join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                                    join SC50 in ProtheusInter.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                                    join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                    join SD20 in ProtheusInter.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                                    where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                                    && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                                    && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                                    && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                                    && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                                    && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                                    && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
                                                    && (int)(object)SD20.D2Emissao >= 20240301
                                                    && SA30.A3Xlogsup != "ANDRE.SALES"
                                                    select new RelatorioAreceberBaixa
                                                    {
                                                        Prefixo = SE50.E5Prefixo,
                                                        Numero = SE50.E5Numero,
                                                        Parcela = SE50.E5Parcela,
                                                        TP = SE50.E5Tipo,
                                                        CliFor = SE50.E5Clifor,
                                                        NomeFor = SA10.A1Nome,
                                                        Naturez = SE50.E5Naturez,
                                                        Vencimento = SE10.E1Vencto,
                                                        Historico = SE50.E5Histor,
                                                        DataBaixa = SE50.E5Data,
                                                        ValorOrig = SE10.E1Valor,
                                                        JurMulta = SE50.E5Vljuros + SE50.E5Vlmulta,
                                                        Correcao = SE50.E5Vlcorre,
                                                        Descon = SE50.E5Vldesco,
                                                        Abatimento = 0,
                                                        Imposto = 0,
                                                        ValorAcess = 0,
                                                        TotalBaixado = SE50.E5Valor,
                                                        Banco = SE50.E5Banco,
                                                        DtDigi = SE50.E5Dtdigit,
                                                        Mot = SE50.E5Motbx,
                                                        Orig = SE50.E5Filorig,
                                                        Vendedor = SC50.C5Nomvend,
                                                        TipoCliente = SA10.A1Clinter,
                                                        CodigoCliente = SA10.A1Xgrinte,
                                                        Login = SA30.A3Xlogin,
                                                        Gestor = SA30.A3Xlogsup,
                                                        DataPedido = SD20.D2Emissao
                                                    }).GroupBy(x => new
                                                    {
                                                        x.Prefixo,
                                                        x.Numero,
                                                        x.Parcela,
                                                        x.TP,
                                                        x.CliFor,
                                                        x.NomeFor,
                                                        x.Naturez,
                                                        x.Vencimento,
                                                        x.Historico,
                                                        x.DataBaixa,
                                                        x.ValorOrig,
                                                        x.JurMulta,
                                                        x.Correcao,
                                                        x.Descon,
                                                        x.Abatimento,
                                                        x.Imposto,
                                                        x.ValorAcess,
                                                        x.TotalBaixado,
                                                        x.Banco,
                                                        x.DtDigi,
                                                        x.Mot,
                                                        x.Orig,
                                                        x.Vendedor,
                                                        x.TipoCliente,
                                                        x.CodigoCliente,
                                                        x.Login,
                                                        x.Gestor,
                                                        x.DataPedido
                                                    }).Select(x => new RelatorioAreceberBaixa
                                                    {
                                                        Prefixo = x.Key.Prefixo,
                                                        Numero = x.Key.Numero,
                                                        Parcela = x.Key.Parcela,
                                                        TP = x.Key.TP,
                                                        CliFor = x.Key.CliFor,
                                                        NomeFor = x.Key.NomeFor,
                                                        Naturez = x.Key.Naturez,
                                                        Vencimento = x.Key.Vencimento,
                                                        Historico = x.Key.Historico,
                                                        DataBaixa = x.Key.DataBaixa,
                                                        ValorOrig = x.Key.ValorOrig,
                                                        JurMulta = x.Key.JurMulta,
                                                        Correcao = x.Key.Correcao,
                                                        Descon = x.Key.Descon,
                                                        Abatimento = x.Key.Abatimento,
                                                        Imposto = x.Key.Imposto,
                                                        ValorAcess = x.Key.ValorAcess,
                                                        TotalBaixado = x.Key.TotalBaixado,
                                                        Banco = x.Key.Banco,
                                                        DtDigi = x.Key.DtDigi,
                                                        Mot = x.Key.Mot,
                                                        Orig = x.Key.Orig,
                                                        Vendedor = x.Key.Vendedor,
                                                        TipoCliente = x.Key.TipoCliente,
                                                        CodigoCliente = x.Key.CodigoCliente,
                                                        Login = x.Key.Login,
                                                        Gestor = x.Key.Gestor,
                                                        DataPedido = x.Key.DataPedido
                                                    }).ToList();

                        #endregion

                        #region Faturamento
                        var query = (from SD20 in ProtheusInter.Sd2010s
                                     join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                     join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                     join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                     where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                     && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                     (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                     && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                     && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
                                     select new
                                     {
                                         Login = SA30.A3Xlogin,
                                         NF = SD20.D2Doc,
                                         Total = SD20.D2Total,
                                         SA10.A1Xgrinte

                                     });

                        var resultado = query.GroupBy(x => new
                        {
                            x.Login,
                            x.NF,
                            x.A1Xgrinte
                        }).Select(x => new RelatorioCirurgiasFaturadas
                        {
                            A3Nome = x.Key.Login,
                            Nf = x.Key.NF,
                            Total = x.Sum(c => c.Total),
                            Linha = x.Key.A1Xgrinte
                        }).ToList();
                        #endregion

                        #region Devolução
                        var teste2 = (from SD10 in ProtheusInter.Sd1010s
                                      join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                      join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                      join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                      join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                      where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                      && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                      && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                      && (int)(object)SD10.D1Dtdigit >= 20200701 && SB10.B1Ugrpint != "082"
                                      && SC50.C5Utpoper == "F" && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                                          SA30.A3Xdescun
                                      }
                             ).GroupBy(x => new
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
                                 x.A3Xdescun
                             });

                        var Relatorio = teste2.Select(x => new RelatorioDevolucaoFat
                        {
                            Filial = x.Key.D1Filial,
                            Clifor = x.Key.D1Fornece,
                            Loja = x.Key.D1Loja,
                            Nome = x.Key.A1Nome,
                            Tipo = x.Key.A1Clinter,
                            Nf = x.Key.D1Doc,
                            Serie = x.Key.D1Serie,
                            Digitacao = x.Key.D1Dtdigit,
                            Total = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc),
                            Valipi = x.Sum(c => c.D1Valipi),
                            Valicm = x.Sum(c => c.D1Valicm),
                            Descon = x.Sum(c => c.D1Valdesc),
                            TotalBrut = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valdesc),
                            A3Nome = x.Key.A3Nome,
                            D1Nfori = x.Key.D1Nfori,
                            D1Seriori = x.Key.D1Seriori,
                            D1Datori = x.Key.D1Datori,
                            Linha = x.Key.A3Xdescun
                        }).ToList();
                        #endregion

                        Equipe.ForEach(x =>
                        {
                            x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").DistinctBy(x => x.Nf).Count();
                            x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total) - Relatorio.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total);

                            x.CirurgiasLicitacoes = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).DistinctBy(x => x.Nf).Count();
                            x.CirurgiasLicitacoesValor = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total) - Relatorio.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total);

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

                    }

                }
                else
                {

                    if (user != "TIAGO.FONSECA")
                    {
                        if (Convert.ToInt32(DataInicio) < 20240228)
                        {
                            #region Gestor

                            if (user == "ARTEMIO.COSTA") user = "LEONARDO.BRITO";

                            var vendedores = ProtheusDenuo.Sa3010s.Where(x => (x.A3Xlogsup == user || x.A3Xlogin == user) && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                            {
                                x.A3Nome,
                                x.A3Xlogin,
                            }).ToList();

                            vendedores.ForEach(x =>
                            {
                                Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                            });

                            #region Valorizado
                            var resultadoValorizado = (from SC50 in ProtheusDenuo.Sc5010s
                                                       join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                       join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                       join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                       join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                       join PA10 in ProtheusDenuo.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                                       from c in Sr.DefaultIfEmpty()
                                                       join SUA10 in ProtheusDenuo.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                                       from a in Rs.DefaultIfEmpty()
                                                       where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                                       && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                       && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user) && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                                x.CirurgiasValorizadasValor = resultadoValorizado.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                            });
                            #endregion

                            #region EmAberto
                            var resultadoEmAberto = (from SC5 in ProtheusDenuo.Sc5010s
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
                                                     && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140" &&
                                                     (SA3.A3Xlogsup == user || SA3.A3Xlogin == user)
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                                x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                            });
                            #endregion

                            #region Faturado

                            var query = (from SD20 in ProtheusDenuo.Sd2010s
                                         join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                         join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                         join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                         && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                         && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                            }).Select(x => new RelatorioCirurgiasFaturadas
                            {
                                A3Nome = x.Key.Login,
                                Nf = x.Key.NF,
                                Total = x.Sum(c => c.Total),
                            }).ToList();
                            #endregion

                            #region Devolucao
                            var teste = (from SD10 in ProtheusInter.Sd1010s
                                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                         join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                         join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                         where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                         && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                         && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                         && (int)(object)SD10.D1Dtdigit >= 20200701 && SC50.C5Utpoper == "F" && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                                             SA10.A1Xgrinte
                                         }
                                         ).ToList();

                            var RelatorioDev = new List<RelatorioDevolucaoFat>();

                            if (teste.Count != 0)
                            {
                                teste.ForEach(x =>
                                {
                                    if (!RelatorioDev.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                                    {

                                        var Iguais = teste
                                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                                        double desconto = 0;
                                        double valipi = 0;
                                        double total = 0;
                                        double valicm = 0;
                                        Iguais.ForEach(x =>
                                        {
                                            desconto += x.D1Valdesc;
                                            valipi += x.D1Valipi;
                                            total += x.D1Total;
                                            valicm += x.D1Valicm;
                                        });

                                        RelatorioDev.Add(new RelatorioDevolucaoFat
                                        {
                                            Filial = x.D1Filial,
                                            Clifor = x.D1Fornece,
                                            Loja = x.D1Loja,
                                            Nome = x.A1Nome,
                                            Tipo = x.A1Clinter,
                                            Nf = x.D1Doc,
                                            Serie = x.D1Serie,
                                            Digitacao = x.D1Dtdigit,
                                            Total = total - desconto,
                                            Valipi = valipi,
                                            Valicm = valicm,
                                            Descon = desconto,
                                            TotalBrut = total - desconto + valipi,
                                            A3Nome = x.A3Nome,
                                            D1Nfori = x.D1Nfori,
                                            D1Seriori = x.D1Seriori,
                                            D1Datori = x.D1Datori,
                                            Linha = x.A1Xgrinte
                                        });
                                    }
                                });
                            }
                            #endregion

                            Equipe.ForEach(x =>
                            {
                                x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User).DistinctBy(x => x.Nf).Count();
                                x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User).Sum(x => x.Total) - RelatorioDev.Where(c => c.Login == x.User).Sum(x => x.Total);

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
                        }
                        else
                        {
                            #region Gestor

                            if (user == "ARTEMIO.COSTA") user = "LEONARDO.BRITO";

                            var vendedores = ProtheusDenuo.Sa3010s.Where(x => (x.A3Xlogsup == user || x.A3Xlogin == user) && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                            {
                                x.A3Nome,
                                x.A3Xlogin,
                            }).ToList();

                            vendedores.ForEach(x =>
                            {
                                Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                            });

                            #region Valorizados
                            var resultadoValorizado = (from SC50 in ProtheusDenuo.Sc5010s
                                                       join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                       join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                       join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                       join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                       join PA10 in ProtheusDenuo.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                                       from c in Sr.DefaultIfEmpty()
                                                       join SUA10 in ProtheusDenuo.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                                       from a in Rs.DefaultIfEmpty()
                                                       where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                                       && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                       && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user) && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                                x.CirurgiasValorizadasValor = resultadoValorizado.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                            });
                            #endregion

                            #region EmAberto

                            var resultadoEmAberto = (from SC5 in ProtheusDenuo.Sc5010s
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
                                                     && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140" &&
                                                     (SA3.A3Xlogsup == user || SA3.A3Xlogin == user)
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                                x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                            });

                            #endregion

                            #region Faturamento

                            var query = (from SD20 in ProtheusDenuo.Sd2010s
                                         join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                         join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                         join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                         && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                         && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
                                         select new
                                         {
                                             Login = SA30.A3Xlogin,
                                             NF = SD20.D2Doc,
                                             Total = SD20.D2Total,
                                             SA10.A1Xgrinte
                                         });

                            var resultado = query.GroupBy(x => new
                            {
                                x.Login,
                                x.NF,
                                x.A1Xgrinte
                            }).Select(x => new RelatorioCirurgiasFaturadas
                            {
                                A3Nome = x.Key.Login,
                                Nf = x.Key.NF,
                                Total = x.Sum(c => c.Total),
                                Linha = x.Key.A1Xgrinte
                            }).ToList();

                            #endregion

                            #region Devolucao
                            var teste = (from SD10 in ProtheusInter.Sd1010s
                                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                         join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                         join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                         where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                         && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                         && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                         && (int)(object)SD10.D1Dtdigit >= 20200701 && SC50.C5Utpoper == "F" && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
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
                                             SA10.A1Xgrinte
                                         }
                                         ).ToList();

                            var RelatorioDev = new List<RelatorioDevolucaoFat>();

                            if (teste.Count != 0)
                            {
                                teste.ForEach(x =>
                                {
                                    if (!RelatorioDev.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                                    {

                                        var Iguais = teste
                                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                                        double desconto = 0;
                                        double valipi = 0;
                                        double total = 0;
                                        double valicm = 0;
                                        Iguais.ForEach(x =>
                                        {
                                            desconto += x.D1Valdesc;
                                            valipi += x.D1Valipi;
                                            total += x.D1Total;
                                            valicm += x.D1Valicm;
                                        });

                                        RelatorioDev.Add(new RelatorioDevolucaoFat
                                        {
                                            Filial = x.D1Filial,
                                            Clifor = x.D1Fornece,
                                            Loja = x.D1Loja,
                                            Nome = x.A1Nome,
                                            Tipo = x.A1Clinter,
                                            Nf = x.D1Doc,
                                            Serie = x.D1Serie,
                                            Digitacao = x.D1Dtdigit,
                                            Total = total - desconto,
                                            Valipi = valipi,
                                            Valicm = valicm,
                                            Descon = desconto,
                                            TotalBrut = total - desconto + valipi,
                                            A3Nome = x.A3Nome,
                                            D1Nfori = x.D1Nfori,
                                            D1Seriori = x.D1Seriori,
                                            D1Datori = x.D1Datori,
                                            Linha = x.A1Xgrinte
                                        });
                                    }
                                });
                            }
                            #endregion

                            #region BaixaLicitacoes

                            var BaixaLicitacoesDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                                        join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                                        equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                                        join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                                        join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                                        join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                        join SD20 in ProtheusDenuo.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                                        where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                                        && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                                        && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                                        && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                                        && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                                        && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                                        && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
                                                        && (int)(object)SD20.D2Emissao >= 20240301
                                                        && (SA30.A3Xlogsup == user || SA30.A3Xlogin == user)
                                                        select new RelatorioAreceberBaixa
                                                        {
                                                            Prefixo = SE50.E5Prefixo,
                                                            Numero = SE50.E5Numero,
                                                            Parcela = SE50.E5Parcela,
                                                            TP = SE50.E5Tipo,
                                                            CliFor = SE50.E5Clifor,
                                                            NomeFor = SA10.A1Nome,
                                                            Naturez = SE50.E5Naturez,
                                                            Vencimento = SE10.E1Vencto,
                                                            Historico = SE50.E5Histor,
                                                            DataBaixa = SE50.E5Data,
                                                            ValorOrig = SE10.E1Valor,
                                                            JurMulta = SE50.E5Vljuros + SE50.E5Vlmulta,
                                                            Correcao = SE50.E5Vlcorre,
                                                            Descon = SE50.E5Vldesco,
                                                            Abatimento = 0,
                                                            Imposto = 0,
                                                            ValorAcess = 0,
                                                            TotalBaixado = SE50.E5Valor,
                                                            Banco = SE50.E5Banco,
                                                            DtDigi = SE50.E5Dtdigit,
                                                            Mot = SE50.E5Motbx,
                                                            Orig = SE50.E5Filorig,
                                                            Vendedor = SC50.C5Nomvend,
                                                            TipoCliente = SA10.A1Clinter,
                                                            CodigoCliente = SA10.A1Xgrinte,
                                                            Login = SA30.A3Xlogin,
                                                            Gestor = SA30.A3Xlogsup,
                                                            DataPedido = SD20.D2Emissao
                                                        }).GroupBy(x => new
                                                        {
                                                            x.Prefixo,
                                                            x.Numero,
                                                            x.Parcela,
                                                            x.TP,
                                                            x.CliFor,
                                                            x.NomeFor,
                                                            x.Naturez,
                                                            x.Vencimento,
                                                            x.Historico,
                                                            x.DataBaixa,
                                                            x.ValorOrig,
                                                            x.JurMulta,
                                                            x.Correcao,
                                                            x.Descon,
                                                            x.Abatimento,
                                                            x.Imposto,
                                                            x.ValorAcess,
                                                            x.TotalBaixado,
                                                            x.Banco,
                                                            x.DtDigi,
                                                            x.Mot,
                                                            x.Orig,
                                                            x.Vendedor,
                                                            x.TipoCliente,
                                                            x.CodigoCliente,
                                                            x.Login,
                                                            x.Gestor,
                                                            x.DataPedido
                                                        }).Select(x => new RelatorioAreceberBaixa
                                                        {
                                                            Prefixo = x.Key.Prefixo,
                                                            Numero = x.Key.Numero,
                                                            Parcela = x.Key.Parcela,
                                                            TP = x.Key.TP,
                                                            CliFor = x.Key.CliFor,
                                                            NomeFor = x.Key.NomeFor,
                                                            Naturez = x.Key.Naturez,
                                                            Vencimento = x.Key.Vencimento,
                                                            Historico = x.Key.Historico,
                                                            DataBaixa = x.Key.DataBaixa,
                                                            ValorOrig = x.Key.ValorOrig,
                                                            JurMulta = x.Key.JurMulta,
                                                            Correcao = x.Key.Correcao,
                                                            Descon = x.Key.Descon,
                                                            Abatimento = x.Key.Abatimento,
                                                            Imposto = x.Key.Imposto,
                                                            ValorAcess = x.Key.ValorAcess,
                                                            TotalBaixado = x.Key.TotalBaixado,
                                                            Banco = x.Key.Banco,
                                                            DtDigi = x.Key.DtDigi,
                                                            Mot = x.Key.Mot,
                                                            Orig = x.Key.Orig,
                                                            Vendedor = x.Key.Vendedor,
                                                            TipoCliente = x.Key.TipoCliente,
                                                            CodigoCliente = x.Key.CodigoCliente,
                                                            Login = x.Key.Login,
                                                            Gestor = x.Key.Gestor,
                                                            DataPedido = x.Key.DataPedido
                                                        }).ToList();

                            #endregion


                            Equipe.ForEach(x =>
                            {
                                x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").DistinctBy(x => x.Nf).Count();
                                x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total) - RelatorioDev.Where(c=> c.Login == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total);

                                x.CirurgiasLicitacoes = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).DistinctBy(x => x.Nf).Count();
                                x.CirurgiasLicitacoesValor = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total) - RelatorioDev.Where(c => c.Login == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total);

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
                        }
                    }
                    else
                    {

                        if (Convert.ToInt32(DataInicio) < 20240228)
                        {
                            #region Tiago

                            var vendedores = ProtheusDenuo.Sa3010s.Where(x => x.A3Xlogsup != "ANDRE.SALES" && x.A3Xlogsup != "" && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                            {
                                x.A3Nome,
                                x.A3Xlogin,
                            }).ToList();

                            vendedores.ForEach(x =>
                            {
                                Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                            });

                            #region Valorizado
                            var resultadoValorizado = (from SC50 in ProtheusDenuo.Sc5010s
                                                       join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                       join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                       join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                       join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                       join PA10 in ProtheusDenuo.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                                       from c in Sr.DefaultIfEmpty()
                                                       join SUA10 in ProtheusDenuo.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                                       from a in Rs.DefaultIfEmpty()
                                                       where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                                       && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                       && SA30.A3Xlogsup != "ANDRE.SALES" && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                                x.CirurgiasValorizadasValor = resultadoValorizado.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                            });
                            #endregion

                            #region EmAberto
                            var resultadoEmAberto = (from SC5 in ProtheusDenuo.Sc5010s
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                                x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                            });
                            #endregion

                            #region Faturamento
                            var query = (from SD20 in ProtheusDenuo.Sd2010s
                                         join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                         join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                         join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                         && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                         && SA30.A3Xlogsup != "ANDRE.SALES"
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
                            }).Select(x => new RelatorioCirurgiasFaturadas
                            {
                                A3Nome = x.Key.Login,
                                Nf = x.Key.NF,
                                Total = x.Sum(c => c.Total),
                            }).ToList();
                            #endregion

                            #region Devolucao
                            var teste = (from SD10 in ProtheusInter.Sd1010s
                                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                         join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                         join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                         where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                         && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                         && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                         && (int)(object)SD10.D1Dtdigit >= 20200701 && SC50.C5Utpoper == "F" && SA30.A3Xlogsup != "ANDRE.SALES"
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
                                             SA10.A1Xgrinte
                                         }
                                         ).ToList();

                            var RelatorioDev = new List<RelatorioDevolucaoFat>();

                            if (teste.Count != 0)
                            {
                                teste.ForEach(x =>
                                {
                                    if (!RelatorioDev.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                                    {

                                        var Iguais = teste
                                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                                        double desconto = 0;
                                        double valipi = 0;
                                        double total = 0;
                                        double valicm = 0;
                                        Iguais.ForEach(x =>
                                        {
                                            desconto += x.D1Valdesc;
                                            valipi += x.D1Valipi;
                                            total += x.D1Total;
                                            valicm += x.D1Valicm;
                                        });

                                        RelatorioDev.Add(new RelatorioDevolucaoFat
                                        {
                                            Filial = x.D1Filial,
                                            Clifor = x.D1Fornece,
                                            Loja = x.D1Loja,
                                            Nome = x.A1Nome,
                                            Tipo = x.A1Clinter,
                                            Nf = x.D1Doc,
                                            Serie = x.D1Serie,
                                            Digitacao = x.D1Dtdigit,
                                            Total = total - desconto,
                                            Valipi = valipi,
                                            Valicm = valicm,
                                            Descon = desconto,
                                            TotalBrut = total - desconto + valipi,
                                            A3Nome = x.A3Nome,
                                            D1Nfori = x.D1Nfori,
                                            D1Seriori = x.D1Seriori,
                                            D1Datori = x.D1Datori,
                                            Linha = x.A1Xgrinte
                                        });
                                    }
                                });
                            }
                            #endregion

                            Equipe.ForEach(x =>
                            {
                                x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User).DistinctBy(x => x.Nf).Count();
                                x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User).Sum(x => x.Total) - RelatorioDev.Where(c => c.A3Nome == x.User).Sum(x => x.Total);

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
                        }
                        else
                        {
                            #region Tiago

                            var vendedores = ProtheusDenuo.Sa3010s.Where(x => x.A3Xlogsup != "ANDRE.SALES" && x.A3Xlogsup != "" && x.DELET != "*" && x.A3Msblql != "1").Select(x => new
                            {
                                x.A3Nome,
                                x.A3Xlogin,
                            }).ToList();

                            vendedores.ForEach(x =>
                            {
                                Equipe.Add(new GestorComercialDash { Nome = x.A3Nome, User = x.A3Xlogin });
                            });

                            #region Valorizado
                            var resultadoValorizado = (from SC50 in ProtheusDenuo.Sc5010s
                                                       join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                                       join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                                       join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                                       join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                       join PA10 in ProtheusDenuo.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                                       from c in Sr.DefaultIfEmpty()
                                                       join SUA10 in ProtheusDenuo.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                                       from a in Rs.DefaultIfEmpty()
                                                       where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                                       && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                                       && SA30.A3Xlogsup != "ANDRE.SALES" && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasValorizadas = resultadoValorizado.Where(c => c.A3XLogin == x.User).DistinctBy(c => c.C5Num).Count();
                                x.CirurgiasValorizadasValor = resultadoValorizado.Where(c => c.A3XLogin == x.User).Sum(x => x.C6Valor);
                            });
                            #endregion

                            #region EmAberto
                            var resultadoEmAberto = (from SC5 in ProtheusDenuo.Sc5010s
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

                            Equipe.ForEach(x =>
                            {
                                x.CirurgiasEmAberto = resultadoEmAberto.Where(c => c.Login == x.User).DistinctBy(x => x.Num).Count();
                                x.CirurgiasEmAbertoValor = resultadoEmAberto.Where(c => c.Login == x.User).Sum(x => x.Valor);
                            });
                            #endregion

                            #region Faturamento
                            var query = (from SD20 in ProtheusDenuo.Sd2010s
                                         join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                         join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                         join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                         && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                         && SA30.A3Xlogsup != "ANDRE.SALES"
                                         select new
                                         {
                                             Login = SA30.A3Xlogin,
                                             NF = SD20.D2Doc,
                                             Total = SD20.D2Total,
                                             SA10.A1Xgrinte
                                         });

                            var resultado = query.GroupBy(x => new
                            {
                                x.Login,
                                x.NF,
                                x.A1Xgrinte
                            }).Select(x => new RelatorioCirurgiasFaturadas
                            {
                                A3Nome = x.Key.Login,
                                Nf = x.Key.NF,
                                Total = x.Sum(c => c.Total),
                                Linha = x.Key.A1Xgrinte
                            }).ToList();
                            #endregion

                            #region Devolucao
                            var teste = (from SD10 in ProtheusInter.Sd1010s
                                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                         join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                         join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                         join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                         join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                         where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                         && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                         && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                         && (int)(object)SD10.D1Dtdigit >= 20200701 && SC50.C5Utpoper == "F" && SA30.A3Xlogsup != "ANDRE.SALES"
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
                                             SA10.A1Xgrinte
                                         }
                                         ).ToList();

                            var RelatorioDev = new List<RelatorioDevolucaoFat>();

                            if (teste.Count != 0)
                            {
                                teste.ForEach(x =>
                                {
                                    if (!RelatorioDev.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                                    {

                                        var Iguais = teste
                                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                                        double desconto = 0;
                                        double valipi = 0;
                                        double total = 0;
                                        double valicm = 0;
                                        Iguais.ForEach(x =>
                                        {
                                            desconto += x.D1Valdesc;
                                            valipi += x.D1Valipi;
                                            total += x.D1Total;
                                            valicm += x.D1Valicm;
                                        });

                                        RelatorioDev.Add(new RelatorioDevolucaoFat
                                        {
                                            Filial = x.D1Filial,
                                            Clifor = x.D1Fornece,
                                            Loja = x.D1Loja,
                                            Nome = x.A1Nome,
                                            Tipo = x.A1Clinter,
                                            Nf = x.D1Doc,
                                            Serie = x.D1Serie,
                                            Digitacao = x.D1Dtdigit,
                                            Total = total - desconto,
                                            Valipi = valipi,
                                            Valicm = valicm,
                                            Descon = desconto,
                                            TotalBrut = total - desconto + valipi,
                                            A3Nome = x.A3Nome,
                                            D1Nfori = x.D1Nfori,
                                            D1Seriori = x.D1Seriori,
                                            D1Datori = x.D1Datori,
                                            Linha = x.A1Xgrinte
                                        });
                                    }
                                });
                            }
                            #endregion

                            #region BaixaLicitacoes

                            var BaixaLicitacoesDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                                        join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                                        equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                                        join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                                        join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                                        join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                                        join SD20 in ProtheusDenuo.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                                        where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                                        && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                                        && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                                        && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                                        && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                                        && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                                        && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
                                                        && (int)(object)SD20.D2Emissao >= 20240301
                                                        && SA30.A3Xlogsup != "ANDRE.SALES"
                                                        select new RelatorioAreceberBaixa
                                                        {
                                                            Prefixo = SE50.E5Prefixo,
                                                            Numero = SE50.E5Numero,
                                                            Parcela = SE50.E5Parcela,
                                                            TP = SE50.E5Tipo,
                                                            CliFor = SE50.E5Clifor,
                                                            NomeFor = SA10.A1Nome,
                                                            Naturez = SE50.E5Naturez,
                                                            Vencimento = SE10.E1Vencto,
                                                            Historico = SE50.E5Histor,
                                                            DataBaixa = SE50.E5Data,
                                                            ValorOrig = SE10.E1Valor,
                                                            JurMulta = SE50.E5Vljuros + SE50.E5Vlmulta,
                                                            Correcao = SE50.E5Vlcorre,
                                                            Descon = SE50.E5Vldesco,
                                                            Abatimento = 0,
                                                            Imposto = 0,
                                                            ValorAcess = 0,
                                                            TotalBaixado = SE50.E5Valor,
                                                            Banco = SE50.E5Banco,
                                                            DtDigi = SE50.E5Dtdigit,
                                                            Mot = SE50.E5Motbx,
                                                            Orig = SE50.E5Filorig,
                                                            Vendedor = SC50.C5Nomvend,
                                                            TipoCliente = SA10.A1Clinter,
                                                            CodigoCliente = SA10.A1Xgrinte,
                                                            Login = SA30.A3Xlogin,
                                                            Gestor = SA30.A3Xlogsup,
                                                            DataPedido = SD20.D2Emissao
                                                        }).GroupBy(x => new
                                                        {
                                                            x.Prefixo,
                                                            x.Numero,
                                                            x.Parcela,
                                                            x.TP,
                                                            x.CliFor,
                                                            x.NomeFor,
                                                            x.Naturez,
                                                            x.Vencimento,
                                                            x.Historico,
                                                            x.DataBaixa,
                                                            x.ValorOrig,
                                                            x.JurMulta,
                                                            x.Correcao,
                                                            x.Descon,
                                                            x.Abatimento,
                                                            x.Imposto,
                                                            x.ValorAcess,
                                                            x.TotalBaixado,
                                                            x.Banco,
                                                            x.DtDigi,
                                                            x.Mot,
                                                            x.Orig,
                                                            x.Vendedor,
                                                            x.TipoCliente,
                                                            x.CodigoCliente,
                                                            x.Login,
                                                            x.Gestor,
                                                            x.DataPedido
                                                        }).Select(x => new RelatorioAreceberBaixa
                                                        {
                                                            Prefixo = x.Key.Prefixo,
                                                            Numero = x.Key.Numero,
                                                            Parcela = x.Key.Parcela,
                                                            TP = x.Key.TP,
                                                            CliFor = x.Key.CliFor,
                                                            NomeFor = x.Key.NomeFor,
                                                            Naturez = x.Key.Naturez,
                                                            Vencimento = x.Key.Vencimento,
                                                            Historico = x.Key.Historico,
                                                            DataBaixa = x.Key.DataBaixa,
                                                            ValorOrig = x.Key.ValorOrig,
                                                            JurMulta = x.Key.JurMulta,
                                                            Correcao = x.Key.Correcao,
                                                            Descon = x.Key.Descon,
                                                            Abatimento = x.Key.Abatimento,
                                                            Imposto = x.Key.Imposto,
                                                            ValorAcess = x.Key.ValorAcess,
                                                            TotalBaixado = x.Key.TotalBaixado,
                                                            Banco = x.Key.Banco,
                                                            DtDigi = x.Key.DtDigi,
                                                            Mot = x.Key.Mot,
                                                            Orig = x.Key.Orig,
                                                            Vendedor = x.Key.Vendedor,
                                                            TipoCliente = x.Key.TipoCliente,
                                                            CodigoCliente = x.Key.CodigoCliente,
                                                            Login = x.Key.Login,
                                                            Gestor = x.Key.Gestor,
                                                            DataPedido = x.Key.DataPedido
                                                        }).ToList();

                            #endregion

                            Equipe.ForEach(x =>
                            {
                                x.FaturadoMes = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").DistinctBy(x => x.Nf).Count();
                                x.FaturadoMesValor = resultado.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total) - RelatorioDev.Where(c => c.A3Nome == x.User && c.Linha != "000011" && c.Linha != "000012").Sum(x => x.Total);

                                x.CirurgiasLicitacoes = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).DistinctBy(x => x.Nf).Count();
                                x.CirurgiasLicitacoesValor = resultado.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total) - RelatorioDev.Where(c => c.A3Nome == x.User && (c.Linha == "000011" || c.Linha == "000012")).Sum(x => x.Total);

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
                        }


                    }

                }
                Equipe = Equipe.OrderBy(x => x.Nome).ToList();
            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex, SGID, "DashBoardGestorComercial", user);

            }
            return Page();
        }

    }
}
