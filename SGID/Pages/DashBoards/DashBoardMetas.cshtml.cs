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
using SGID.Models.RH;

namespace SGID.Pages.DashBoards
{
    [Authorize]
    public class DashBoardMetasModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public List<ValoresEmAberto> ValoresEmAberto { get; set; } = new List<ValoresEmAberto>();

        public List<ValoresEmAberto> Faturamento { get; set; } = new List<ValoresEmAberto>();

        public string Ano { get; set; }

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
                                 (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains(SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                 && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                                 select new
                                 {
                                     Login = SA30.A3Xlogin,
                                     NF = SD20.D2Doc,
                                     Total = SD20.D2Total,
                                     TotalBrut = SD20.D2Valbrut,
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
                        TotalBrut = x.Sum(c=>c.TotalBrut),
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
                                          TotalBrut = SD20.D2Valbrut,
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
                        TotalBrut = x.Sum(c => c.TotalBrut),
                        Linha = x.Key.Linha
                    }).ToList();
                    #endregion
                
                var resultado = resultadoDenuo.Concat(resultadoInter).ToList();
                
                
                var resultado2 = resultado.GroupBy(x => new { x.Linha }).Select(x => new
                    {
                        Linha = x.Key.Linha,
                        Quant = x.Count(),
                        Valor = x.Sum(c => c.Total),
                        ValorBrut = x.Sum(c=> c.TotalBrut)
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
                        new RelatorioFaturamentoLinhas { Nome = "QTDA. CIRURGIAS", Quant = resultado2.Sum(x => x.Quant) }
                    };
                
                LinhasValor.ForEach(x =>
                {
                    LinhasTela.Add(new RelatorioFaturamentoLinhas { Nome = x.Nome, Quant = x.Quant });
                });
                
                LinhasTela.Add(new RelatorioFaturamentoLinhas { Nome = "VALOR CIRURGIAS", Quant = resultado2.Sum(x => x.Valor) });
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

                Faturamento.Add(new ValoresEmAberto { Nome = "Sub Inter", Valor = FaturadoSubInter.Sum(c => c.Total) });

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

                Faturamento.Add(new ValoresEmAberto { Nome = "Sub Denuo", Valor = FaturadoSubDenuo.Sum(c => c.Total) });

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
                Logger.Log(ex, SGID, "DashBoardMetas Faturados", user);

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

                if (Mes == "13")
                {
                    DataInicio = $"{Ano}0101";

                    DataFim = $"{Ano}1231";



                }
                else
                {
                    DataInicio = $"{Ano}{Mes}01";

                    DataFim = $"{Ano}{Mes}31";
                }


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
                                              && SC6.C6Nota == "" && SC6.C6Blq != "R" && SC6.DELET != "*"
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

                ValoresEmAberto.Add(new ValoresEmAberto { Nome = "Sub Inter", Valor = SubdistribuidorInter });
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

                ValoresEmAberto.Add(new ValoresEmAberto { Nome = "Sub Denuo", Valor = SubdistribuidorDenuo });
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
