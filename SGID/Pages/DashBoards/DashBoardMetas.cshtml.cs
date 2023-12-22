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
using SGID.Data.ViewModel;
using SGID.Models.Financeiro;
using SGID.Models.Controladoria.FaturamentoNF;

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

        #region Comissoes

        public List<Time> UsersComercial { get; set; }
        public List<TimeRH> UsuariosComercial { get; set; } = new List<TimeRH>();

        public List<TimeRH> ComercialOrtopedia { get; set; } = new List<TimeRH>();
        public List<TimeRH> ComercialBuco { get; set; } = new List<TimeRH>();
        public List<TimeRH> ComercialTorax { get; set; } = new List<TimeRH>();
        public List<TimeRH> ComercialInterior { get; set; } = new List<TimeRH>();


        public List<TimeRH> UsuariosLicitacoes { get; set; } = new List<TimeRH>();
        public List<TimeRH> UsuariosSub { get; set; } = new List<TimeRH>();
        public List<TimeADM> Users { get; set; }
        public List<TimeADMRH> Usuarios { get; set; } = new List<TimeADMRH>();

        public List<TimeDental> UsersDental { get; set; }
        public List<TimeDentalRH> UsuariosDental { get; set; } = new List<TimeDentalRH>();

        #endregion

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
                                        }).ToList();
                var dolar = ProtheusInter.Sm2010s.Where(x => x.DELET != "*" && x.M2Moeda2 != 0).OrderByDescending(x => x.M2Data).FirstOrDefault();
                Faturamento.Add(new ValoresEmAberto { Nome = "SUB INTER", Valor = FaturadoSubInter.Sum(c => c.Total) });
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

                                        }).ToList();

                var dolar2 = ProtheusDenuo.Sm2010s.Where(x => x.DELET != "*" && x.M2Moeda2 != 0).OrderByDescending(x => x.M2Data).FirstOrDefault();

                Faturamento.Add(new ValoresEmAberto { Nome = "SUB DENUO", Valor = FaturadoSubDenuo.Sum(c => c.Total) });

                #endregion*/


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
                    Quant = x.GroupBy(c=> c.Nf).Count(),
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

                    if(valor2 != null) 
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

                var dolar = ProtheusInter.Sm2010s.Where(x => x.DELET != "*" && x.M2Moeda2 != 0).OrderByDescending(x => x.M2Data).FirstOrDefault();

                ValoresEmAberto.Add(new ValoresEmAberto { Nome = "SUB INTER", Valor = SubdistribuidorInter * dolar.M2Moeda2 });
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

                var dolar2 = ProtheusDenuo.Sm2010s.Where(x => x.DELET != "*" && x.M2Moeda2 != 0).OrderByDescending(x => x.M2Data).FirstOrDefault();

                ValoresEmAberto.Add(new ValoresEmAberto { Nome = "SUB DENUO", Valor = SubdistribuidorDenuo * dolar2.M2Moeda2 });
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

        public JsonResult OnPostVencidos()
        {
            string user = User.Identity.Name.Split("@")[0].ToUpper();
            DateTime data = DateTime.Now.AddMonths(9);
            string mesInicio = data.Month.ToString("D2");
            string anoInicio = (data.Year - 1).ToString();
            string DataInicio = $"{data.Year - 1}{data.Month:D2}{data.Day:D2}";

            var NaoBaixados = (from SE10 in ProtheusDenuo.Se1010s
                               join SA10 in ProtheusDenuo.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
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
                            }).OrderByDescending(x => x.Valor).ToList();

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
            string DataInicio = $"{data.Year - 1}{data.Month:D2}{data.Day:D2}";

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
                            }).OrderByDescending(x => x.Valor).ToList();

            var valores = new
            {
                Valores = NaoBaixados,
                ValorTotal = NaoBaixados.Sum(x => x.Valor)
            };

            return new JsonResult(valores);
        }

        public JsonResult OnPostComissoes(string Mes, string Ano)
        {
            #region Dados

            string DataInicio = "";
            string DataFim = "";
            if (Mes != "13")
            {
                string Tempo = $"{Mes}/01/{Ano}";

                var date = DateTime.Parse(Tempo);

                //MesAno = date.ToString("MM").ToUpper();
                //this.Ano = date.ToString("yyyy").ToUpper();

                string data = date.ToString("yyyy/MM").Replace("/", "");
                DataInicio = data + "01";
                DataFim = data + "31";
            }
            else
            {
                DataInicio = $"{Ano}0101";
                DataFim = $"{Ano}1231"; ;
            }

            int[] CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };

            var cf = new[] { "5551", "6551", "6107", "6109" };

            #endregion

            #region Dental

                UsersDental = SGID.TimeDentals.OrderBy(x => x.Integrante).ToList();

                #region Faturado
                var query3 = (from SD20 in ProtheusDenuo.Sd2010s
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
                              || cf.Contains(SD20.D2Cf)) && (int)(object)SD20.D2Emissao >= (int)(object)DataInicio
                              && (int)(object)SD20.D2Emissao <= (int)(object)DataFim
                              && SD20.D2Quant != 0 && SC50.C5Xtipopv == "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140"
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
                                  D2_DESCON = SD20.D2Descon,
                                  A3_LOGIN = c.A3Xlogin
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
                                  Prod.B1_DESC,
                                  Prod.A3_LOGIN
                              });

                var Faturamento = query3.Select(x => new RelatorioFaturamentoDental
                {
                    D2_FILIAL = x.Key.D2_FILIAL,
                    D2_CLIENTE = x.Key.D2_CLIENTE,
                    D2_LOJA = x.Key.D2_LOJA,
                    A1_NOME = x.Key.A1_NOME,
                    A1_CLINTER = x.Key.A1_CLINTER,
                    D2_DOC = x.Key.D2_DOC,
                    D2_SERIE = x.Key.D2_SERIE,
                    D2_EMISSAO = $"{x.Key.D2_EMISSAO.Substring(6, 2)}/{x.Key.D2_EMISSAO.Substring(4, 2)}/{x.Key.D2_EMISSAO.Substring(0, 4)}",
                    D2_PEDIDO = x.Key.D2_PEDIDO,
                    C5_UNUMAGE = x.Key.C5_UNUMAGE,
                    C5_EMISSAO = x.Key.C5_EMISSAO,
                    C5_VEND1 = x.Key.C5_VEND1,
                    A3_NOME = x.Key.A3_NOME,
                    C5_X_DTCIR = x.Key.C5_X_DTCIR,
                    C5_X_NMMED = x.Key.C5_X_NMMED,
                    C5_X_NMPAC = x.Key.C5_X_NMPAC,
                    C5_X_NMPLA = x.Key.C5_X_NMPLA,
                    C5_UTPOPER = x.Key.C5_UTPOPER,
                    C6_PRODUTO = x.Key.C6_PRODUTO,
                    B1_DESC = x.Key.B1_DESC,
                    D2_TOTAL = x.Sum(c => c.D2_TOTAL),
                    D2_QUANT = x.Sum(c => c.D2_QUANT),
                    D2_VALIPI = x.Sum(c => c.D2_VALIPI),
                    D2_VALICM = x.Sum(c => c.D2_VALICM),
                    D2_DESCON = x.Sum(c => c.D2_DESCON),
                    A3_LOGIN = x.Key.A3_LOGIN
                }).OrderBy(x => x.A3_NOME).ToList();
                #endregion

                #region Devolucao
                var CfDevolucao = new string[] { "1202", "2202", "3202", "1553", "2553" };

                var queryDevolucao3 = (from SD10 in ProtheusDenuo.Sd1010s
                                       join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Fornece = SF20.F2Cliente, Loja = SF20.F2Loja }
                                       join SD20 in ProtheusDenuo.Sd2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja, Item = SD10.D1Itemori } equals new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Fornece = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Item }
                                       join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                       join SA10 in ProtheusDenuo.Sa1010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA10.A1Cod, Loja = SA10.A1Loja }
                                       join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
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
                                           SA10.A1Mun,
                                           SA30.A3Xlogin
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
                                              x.C5Nomvend,
                                              x.A3Xlogin
                                          });

                var Devolucao = queryDevolucao3.Select(x => new RelatorioDevolucaoDental
                {
                    Emissao = $"{x.Key.D1Emissao.Substring(6, 2)}/{x.Key.D1Emissao.Substring(4, 2)}/{x.Key.D1Emissao.Substring(0, 4)}",
                    Filial = x.Key.D1Filial,
                    NumNF = x.Key.D1Doc,
                    EmissaoOrig = $"{x.Key.D2Emissao.Substring(6, 2)}/{x.Key.D2Emissao.Substring(4, 2)}/{x.Key.D2Emissao.Substring(0, 4)}",
                    NFOrig = x.Key.D2Doc,
                    CodVend = x.Key.C5Vend1,
                    Vendedor = x.Key.C5Nomvend,
                    Cliente = x.Key.A1Nome,
                    Total = -(x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valipi)),
                    Login = x.Key.A3Xlogin
                }).OrderBy(x => x.Vendedor).ToList();

                #endregion

                UsersDental.ForEach(x =>
                {
                    var usuario = x.Integrante.ToUpper();


                    var time = new TimeDentalRH
                    {
                        User = x,
                    };


                    if (usuario != "MARCOS.PARRA")
                    {
                        time.Faturado += Faturamento.Where(x => x.A3_LOGIN.Trim() == usuario).Sum(x => x.D2_TOTAL) + Devolucao.Where(x => x.Login.Trim() == usuario).Sum(x => x.Total);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                        time.Meta = time.User.Meta / 12;
                        time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);
                        time.Salario = time.User.Salario;

                        time.Total = time.Salario + time.Comissao;
                        UsuariosDental.Add(time);
                    }
                    else
                    {
                        time.FaturadoEquipe += Faturamento.Sum(x => x.D2_TOTAL) + Devolucao.Sum(x => x.Total);

                        time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                        time.Meta = time.User.Meta / 12;
                        time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);
                        time.Salario = time.User.Salario;

                        time.Total = time.Salario + time.Comissao;
                        UsuariosDental.Add(time);
                    }
                });

                #endregion

            #region Comercial

                #region INTERMEDIC

                #region Faturado
                var query = (from SD20 in ProtheusInter.Sd2010s
                             join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             join SC60 in ProtheusInter.Sc6010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Pedido = SC60.C6Num, Cliente = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                             where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SC60.DELET != "*"
                             && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                             ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                             && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
                             select new
                             {
                                 NF = SD20.D2Doc,
                                 Total = SD20.D2Total,
                                 Data = SD20.D2Emissao,
                                 Login = SA30.A3Xlogin,
                                 Gestor = SA30.A3Xlogsup,
                                 Linha = SA30.A3Xdescun,
                                 DOR = SB10.B1Ugrpint,
                                 Codigo = SA10.A1Xgrinte
                             });

                var resultadoInter = query.GroupBy(x => new
                {
                    x.NF,
                    x.Data,
                    x.Login,
                    x.Gestor,
                    x.Linha,
                    x.DOR,
                    x.Codigo
                }).Select(x => new
                {
                    Nf = x.Key.NF,
                    Total = x.Sum(c => c.Total),
                    Data = x.Key.Data,
                    Login = x.Key.Login.Trim(),
                    Gestor = x.Key.Gestor.Trim(),
                    Linha = x.Key.Linha.Trim(),
                    DOR = x.Key.DOR,
                    Codigo = x.Key.Codigo
                }).ToList();
                #endregion

                #region Devolucao

                var DevolucaoInter = new List<RelatorioDevolucaoFat>();

                var teste = (from SD10 in ProtheusInter.Sd1010s
                             join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                             join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                             join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                             where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                              && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                              && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                              && (int)(object)SD10.D1Dtdigit >= 20200701
                              && SC50.C5Utpoper == "F"
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
                                 SA30.A3Xlogin,
                                 SA30.A3Xlogsup,
                                 SA30.A3Xdescun,
                                 DOR = SB10.B1Ugrpint,
                                 SA10.A1Xgrinte,
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
                                 x.A3Xdescun,
                                 x.A1Xgrinte,
                                 x.A3Xlogin,
                                 x.A3Xlogsup,
                                 x.DOR
                             });

                DevolucaoInter = teste.Select(x => new RelatorioDevolucaoFat
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
                    Linha = x.Key.A3Xdescun.Trim(),
                    Codigo = x.Key.A1Xgrinte,
                    DOR = x.Key.DOR,
                    Login = x.Key.A3Xlogin.Trim(),
                    Gestor = x.Key.A3Xlogsup.Trim(),
                }).ToList();

                #endregion

                #region Baixa

                var BaixaInter = (from SE50 in ProtheusInter.Se5010s
                                  join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                  equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                  join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                  join SC50 in ProtheusInter.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                  join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                  where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                  && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                  && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                  && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                  && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                  && (int)(object)SE50.E5Data <= (int)(object)DataFim
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
                                      Vendedor = SA30.A3Nome,
                                      TipoCliente = SA10.A1Clinter,
                                      CodigoCliente = SA10.A1Xgrinte,
                                      Login = SA30.A3Xlogin,
                                      Gestor = SA30.A3Xlogsup
                                  }).ToList();
                #endregion

                #endregion

                #region DENUO

                #region Faturado

                var query2 = (from SD20 in ProtheusDenuo.Sd2010s
                              join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                              join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                              join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                              join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                              join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                              where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                 && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                 (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                 && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SC50.C5Xtipopv != "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
                              select new
                              {
                                  NF = SD20.D2Doc,
                                  Total = SD20.D2Total,
                                  Login = SA30.A3Xlogin,
                                  Data = SD20.D2Emissao,
                                  Gestor = SA30.A3Xlogsup,
                                  Linha = SA30.A3Xdescun,
                                  Codigo = SA10.A1Xgrinte
                              });

                var resultadoDenuo = query2.GroupBy(x => new
                {
                    x.NF,
                    x.Login,
                    x.Gestor,
                    x.Data,
                    x.Linha,
                    x.Codigo
                }).Select(x => new
                {
                    Nf = x.Key.NF,
                    Total = x.Sum(c => c.Total),
                    Login = x.Key.Login.Trim(),
                    Data = x.Key.Data,
                    Gestor = x.Key.Gestor.Trim(),
                    Linha = x.Key.Linha.Trim(),
                    Codigo = x.Key.Codigo
                }).ToList();
                #endregion

                #region Devolucao
                var DevolucaoDenuo = new List<RelatorioDevolucaoFat>();

                var teste2 = (from SD10 in ProtheusDenuo.Sd1010s
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
                                  SA30.A3Xlogin,
                                  SA30.A3Xlogsup,
                                  SD10.D1Nfori,
                                  SD10.D1Seriori,
                                  SD10.D1Datori,
                                  SD10.D1Emissao,
                                  SA30.A3Xdescun,
                                  SA10.A1Xgrinte,
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
                                 x.A3Xdescun,
                                 x.A1Xgrinte,
                                 x.A3Xlogin,
                                 x.A3Xlogsup
                             });

                DevolucaoDenuo = teste2.Select(x => new RelatorioDevolucaoFat
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
                    Login = x.Key.A3Xlogin.Trim(),
                    Linha = x.Key.A3Xdescun.Trim(),
                    Gestor = x.Key.A3Xlogsup.Trim(),
                    Codigo = x.Key.A1Xgrinte
                }).ToList();
                #endregion

                #region Baixa

                var BaixaDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                  join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                  equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                  join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                  join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num
                                  join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                  where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                  && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                  && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                  && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                  && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                  && (int)(object)SE50.E5Data <= (int)(object)DataFim
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
                                      Vendedor = SA30.A3Nome,
                                      TipoCliente = SA10.A1Clinter,
                                      CodigoCliente = SA10.A1Xgrinte,
                                      Login = SA30.A3Xlogin,
                                      Gestor = SA30.A3Xlogsup
                                  }).ToList();

                #endregion

                #endregion

                UsersComercial = SGID.Times.Where(x => x.Status).OrderBy(x => x.Integrante).ToList();

                UsersComercial.ForEach(x =>
                {
                    var usuario = x.Integrante.ToUpper();

                    var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == x.Id).ToList();

                    var time = new TimeRH
                    {
                        User = x,
                    };

                    if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                    {
                        time.Linha = resultadoDenuo.FirstOrDefault(x => x.Login == usuario)?.Linha;
                        if (time.User.Integrante.ToUpper() == "EDUARDO.ARONI")
                        {
                            time.Linha = "ORTOPEDIA";
                        }
                        else if (time.User.Integrante.ToUpper() == "TIAGO.FONSECA" || time.User.Integrante.ToUpper() == "ARTEMIO.COSTA")
                        {
                            time.Linha = "BUCOMAXILO";
                        }
                        else if (time.Linha == null)
                        {
                            time.Linha = resultadoInter.FirstOrDefault(x => x.Login == usuario)?.Linha;
                        }

                        time.Faturado += resultadoInter.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                        time.Faturado += resultadoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);
                        if (usuario == "MICHEL.SAMPAIO")
                        {
                            time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                            time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total);
                        }
                        else
                        {
                            time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total);
                            time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario).Sum(x => x.Total);
                        }
                        int i = 0;
                        Produtos.ForEach(prod =>
                        {
                            if (usuario == "TIAGO.FONSECA")
                            {
                                if (i == 0)
                                {
                                    time.FaturadoEquipe = 0;
                                    time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                    time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() ).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() ).Sum(x => x.Total);
                                    i++;
                                }
                                else
                                {
                                    time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                    time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                }
                            }
                            else
                            {
                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                            }
                        });
                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                        time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                        time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);

                        time.Salario = time.User.Salario;

                        time.Meta = time.User.Meta / 12;
                        time.MetaAtingimento = (((time.Faturado + time.FaturadoProduto) / (time.User.Meta / 12)) * 100);

                        time.Total = time.Salario + time.Comissao + time.ComissaoEquipe + time.ComissaoProduto;


                        if (time.User.TipoVendedor == "INTERIOR")
                        {
                            ComercialInterior.Add(time);
                        }
                        else if (time.Linha == "BUCOMAXILO")
                        {
                            ComercialBuco.Add(time);
                        }
                        else if (time.Linha == "TORAX")
                        {
                            ComercialTorax.Add(time);
                        }
                        else
                        {
                            ComercialOrtopedia.Add(time);
                        }


                    }
                    else if (x.TipoFaturamento != "S")
                    {


                        time.Faturado = BaixaDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                        time.Salario = time.User.Salario;
                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                        time.Meta = time.User.Meta / 12;
                        time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                        time.Total = time.Salario + time.Comissao;

                        UsuariosLicitacoes.Add(time);
                    }
                    else
                    {

                        time.Faturado = BaixaDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                        time.Salario = time.User.Salario;
                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                        time.Meta = time.User.Meta / 12;
                        time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                        time.Total = time.Salario + time.Comissao;

                        UsuariosSub.Add(time);
                    }
                });

                #endregion

            #region ADM

                Users = SGID.TimeADMs.OrderBy(x => x.Integrante).ToList();

                Users.ForEach(x =>
                {
                    var usuario = x.Integrante.ToUpper();


                    var time = new TimeADMRH
                    {
                        User = x,
                    };

                    time.Faturado += resultadoInter.Sum(x => x.Total) - DevolucaoInter.Sum(x => x.Total);
                    time.Faturado += resultadoDenuo.Sum(x => x.Total) - DevolucaoDenuo.Sum(x => x.Total);

                    time.Salario = time.User.Salario;

                    time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                    time.Meta = 5200000;
                    time.MetaAtingimento = ((time.Faturado / time.Meta) * 100);


                    time.Total = time.Salario + time.Comissao;
                    Usuarios.Add(time);
                });

            #endregion                

            var Salario = ComercialOrtopedia.Sum(x => x.Salario) + ComercialBuco.Sum(x => x.Salario) + ComercialTorax.Sum(x => x.Salario) + ComercialInterior.Sum(x => x.Salario) + UsuariosLicitacoes.Sum(x => x.Salario) + UsuariosSub.Sum(x => x.Salario) + Usuarios.Sum(x => x.Salario) + UsuariosDental.Sum(x => x.Salario);

            var Comissoes = UsuariosLicitacoes.Sum(x => x.Comissao) + UsuariosSub.Sum(x => x.Comissao) + Usuarios.Sum(x => x.Comissao) + ComercialOrtopedia.Sum(x => x.Comissao)
                + ComercialBuco.Sum(x => x.Comissao) + ComercialTorax.Sum(x => x.Comissao) + ComercialInterior.Sum(x => x.Comissao) + UsuariosDental.Sum(x => x.Comissao)
                + ComercialOrtopedia.Sum(x => x.ComissaoEquipe) + ComercialBuco.Sum(x => x.ComissaoEquipe) + ComercialTorax.Sum(x => x.ComissaoEquipe) + ComercialInterior.Sum(x => x.ComissaoEquipe)
                + UsuariosDental.Sum(x => x.ComissaoEquipe) + ComercialOrtopedia.Sum(x => x.ComissaoProduto) + ComercialBuco.Sum(x => x.ComissaoProduto) + ComercialTorax.Sum(x => x.ComissaoProduto) + ComercialInterior.Sum(x => x.ComissaoProduto);

            if (Mes != "13")
            {
                return new JsonResult(new
                {
                    Salario = Salario,
                    Comissoes = Comissoes
                });
            }
            else
            {
                return new JsonResult(new
                {
                    Salario = Salario * 12,
                    Comissoes = Comissoes
                });
            }
        }
    }
}