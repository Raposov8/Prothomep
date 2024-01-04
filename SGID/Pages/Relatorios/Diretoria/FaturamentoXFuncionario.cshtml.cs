using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models;
using SGID.Models.Account.RH;
using SGID.Models.Denuo;
using SGID.Models.Financeiro;
using SGID.Models.Inter;
using SGID.Models.RH;

namespace SGID.Pages.Relatorios.Diretoria
{
    [Authorize]
    public class FaturamentoXFuncionarioModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

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

        public string Mes { get; set; }
        public string Ano { get; set; }


        public FaturamentoXFuncionarioModel(ApplicationDbContext sgid,TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            SGID = sgid;
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
        }

        public void OnGet()
        {
        }

        public IActionResult OnGetDiretoria(string Mes, string Ano, string Empresa)
        {
            #region Dados
            string Tempo = $"{Mes}/01/{Ano}";

            var date = DateTime.Parse(Tempo);

            //MesAno = date.ToString("MM").ToUpper();
            //this.Ano = date.ToString("yyyy").ToUpper();

            string data = date.ToString("yyyy/MM").Replace("/", "");
            string DataInicio = data + "01";
            string DataFim = data + "31";
            int[] CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };

            var dataini = Convert.ToInt32(DataInicio);
            var cf = new[] { "5551", "6551", "6107", "6109" };

            this.Mes = Mes;
            this.Ano = Ano;

            #endregion

            if (Empresa != "")
            {
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
                        time.Garantia = time.User.Garantia;

                        time.Total = time.Salario + time.Garantia + time.Comissao;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }

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

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }

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

                #region BaixaSub

                var BaixaSubInter = (from SE50 in ProtheusInter.Se5010s
                                     join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                     equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                     join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                     join SC50 in ProtheusInter.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                     from c in sr.DefaultIfEmpty()
                                     where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                     && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                     && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                     && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                     && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                     && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                     && SA10.A1Clinter == "S"
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
                                         Vendedor = c.C5Nomvend,
                                         TipoCliente = SA10.A1Clinter,
                                         Empresa = "INTERMEDIC"
                                     }
                             ).ToList();
                #endregion

                #region BaixaLicitacoes

                var BaixaLicitacoesInter = (from SE50 in ProtheusInter.Se5010s
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
                                            && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                DataPedido = SC50.C5Emissao
                                            }).ToList();

                #endregion

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

                #region BaixaSub

                var BaixaSubDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                     join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                     equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                     join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                     join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                     from c in sr.DefaultIfEmpty()
                                     where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                     && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                     && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                     && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                     && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                     && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                     && SA10.A1Clinter == "S"
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
                                         Vendedor = c.C5Nomvend,
                                         TipoCliente = SA10.A1Clinter,
                                         Empresa = "DENUO"
                                     }
                             ).ToList();

                #endregion

                #region BaixaLicitacoes

                var BaixaLicitacoesDenuo = (from SE50 in ProtheusDenuo.Se5010s
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
                                            && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                DataPedido = SC50.C5Emissao,
                                                Empresa = "DENUO"
                                            }).ToList();

                #endregion

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

                        if (dataini <=  20231231)
                        {

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
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                        i++;
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                    }
                                }
                                else
                                {
                                    time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                    time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                }
                            });
                        }
                        else
                        {
                            time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                            time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                            time.Faturado += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                            time.Faturado += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                            if (usuario == "MICHEL.SAMPAIO")
                            {
                                time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                                time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                            }
                            else
                            {
                                time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                            }
                            int i = 0;
                            Produtos.ForEach(prod =>
                            {
                                if (usuario == "TIAGO.FONSECA")
                                {
                                    if (i == 0)
                                    {
                                        time.FaturadoEquipe = 0;
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        i++;
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                    }
                                }
                                else
                                {
                                    time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                    time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                }
                            });

                        }

                        

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                        time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                        time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);

                        time.Salario = time.User.Salario;
                        time.Meta = time.User.Meta / 12;
                        time.MetaAtingimento = (((time.Faturado + time.FaturadoProduto) / (time.User.Meta / 12)) * 100);

                        time.Teto = time.User.Teto;
                        time.Garantia = time.User.Garantia;

                        time.Total = time.Salario + time.Garantia + time.Comissao + time.ComissaoEquipe + time.ComissaoProduto;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }


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


                        time.Faturado = BaixaLicitacoesDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaLicitacoesInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                        time.Salario = time.User.Salario;
                        time.Meta = time.User.Meta / 12;
                        time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                        time.Teto = time.User.Teto;

                        time.Garantia = time.User.Garantia;

                        time.Total = time.Salario + time.Garantia + time.Comissao;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }

                        UsuariosLicitacoes.Add(time);
                    }
                    else
                    {

                        time.Faturado = BaixaSubDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaSubInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                        time.Salario = time.User.Salario;
                        time.Meta = time.User.Meta / 12;
                        time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                        time.Teto = time.User.Teto;

                        time.Garantia = time.User.Garantia;

                        time.Total = time.Salario + time.Garantia + time.Comissao;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }

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

                    time.Teto = time.User.Teto;

                    time.Total = time.Salario + time.Comissao;

                    if (time.Teto > 0 && time.Total > time.Teto)
                    {
                        time.Paga = time.Teto;
                    }
                    else
                    {
                        time.Paga = time.Total;
                    }

                    Usuarios.Add(time);
                });

                #endregion
            }
            else
            {
                if (Empresa == "INTERMEDIC")
                {
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

                    #region BaixaSub

                    var BaixaSubInter = (from SE50 in ProtheusInter.Se5010s
                                         join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                         equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                         join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                         join SC50 in ProtheusInter.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                         from c in sr.DefaultIfEmpty()
                                         where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                         && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                         && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                         && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                         && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                         && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                         && SA10.A1Clinter == "S"
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
                                             Vendedor = c.C5Nomvend,
                                             TipoCliente = SA10.A1Clinter,
                                             Empresa = "INTERMEDIC"
                                         }
                                 ).ToList();
                    #endregion

                    #region BaixaLicitacoes

                    var BaixaLicitacoesInter = (from SE50 in ProtheusInter.Se5010s
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
                                                && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                    DataPedido = SC50.C5Emissao
                                                }).ToList();

                    #endregion

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
                            time.Linha = resultadoInter.FirstOrDefault(x => x.Login == usuario)?.Linha;
                            if (time.User.Integrante.ToUpper() == "EDUARDO.ARONI")
                            {
                                time.Linha = "ORTOPEDIA";
                            }
                            else if (time.User.Integrante.ToUpper() == "TIAGO.FONSECA" || time.User.Integrante.ToUpper() == "ARTEMIO.COSTA")
                            {
                                time.Linha = "BUCOMAXILO";
                            }

                            if (dataini <=  20231231)
                            {

                                time.Faturado += resultadoInter.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                                }
                                else
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total);
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
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                    }
                                });
                            }
                            else
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.Faturado += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                                }
                                else
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                }
                                int i = 0;
                                Produtos.ForEach(prod =>
                                {
                                    if (usuario == "TIAGO.FONSECA")
                                    {
                                        if (i == 0)
                                        {
                                            time.FaturadoEquipe = 0;
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                    }
                                });

                            }
                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                            time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = (((time.Faturado + time.FaturadoProduto) / (time.User.Meta / 12)) * 100);
                            time.Garantia = time.User.Garantia;

                            time.Teto = time.User.Teto;

                            time.Total = time.Salario + time.Garantia + time.Comissao + time.ComissaoEquipe + time.ComissaoProduto;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

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

                            time.Faturado = BaixaLicitacoesInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            time.Salario = time.User.Salario;

                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                            time.Teto = time.User.Teto;
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

                            UsuariosLicitacoes.Add(time);
                        }
                        else
                        {

                            time.Faturado = BaixaSubInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                            time.Teto = time.User.Teto;
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

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

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                        time.Salario = time.User.Salario;

                        time.Meta = 2600000;
                        time.MetaAtingimento = ((time.Faturado / time.Meta) * 100);

                        time.Teto = time.User.Teto;

                        time.Total = time.Salario + time.Comissao;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }
                        Usuarios.Add(time);
                    });

                    #endregion
                }
                else
                {
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
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;
                            UsuariosDental.Add(time);
                        }
                        else
                        {
                            time.FaturadoEquipe += Faturamento.Sum(x => x.D2_TOTAL) + Devolucao.Sum(x => x.Total);

                            time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);
                            time.Salario = time.User.Salario;
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

                            UsuariosDental.Add(time);
                        }
                    });

                    #endregion

                    #region Comercial

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

                    #region BaixaSub

                    var BaixaSubDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                         join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                         equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                         join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                         join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                         from c in sr.DefaultIfEmpty()
                                         where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                         && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                         && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                         && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                         && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                         && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                         && SA10.A1Clinter == "S"
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
                                             Vendedor = c.C5Nomvend,
                                             TipoCliente = SA10.A1Clinter,
                                             Empresa = "DENUO"
                                         }
                                 ).ToList();

                    #endregion

                    #region BaixaLicitacoes

                    var BaixaLicitacoesDenuo = (from SE50 in ProtheusDenuo.Se5010s
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
                                                && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                    DataPedido = SC50.C5Emissao,
                                                    Empresa = "DENUO"
                                                }).ToList();

                    #endregion

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

                            if (dataini <=  20231231)
                            {

                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);
                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total);
                                }
                                else
                                {
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
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                    }
                                });
                            }
                            else
                            {
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.Faturado += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                                }
                                else
                                {
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                }
                                int i = 0;
                                Produtos.ForEach(prod =>
                                {
                                    if (usuario == "TIAGO.FONSECA")
                                    {
                                        if (i == 0)
                                        {
                                            time.FaturadoEquipe = 0;
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                    }
                                });

                            }
                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                            time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = (((time.Faturado + time.FaturadoProduto) / (time.User.Meta / 12)) * 100);
                            time.Garantia = time.User.Garantia;
                            
                            time.Total = time.Salario + time.Comissao + time.ComissaoEquipe + time.ComissaoProduto;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

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


                            time.Faturado = BaixaLicitacoesDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

                            UsuariosLicitacoes.Add(time);
                        }
                        else
                        {

                            time.Faturado = BaixaSubDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

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

                        time.Faturado += resultadoDenuo.Sum(x => x.Total) - DevolucaoDenuo.Sum(x => x.Total);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                        time.Salario = time.User.Salario;

                        time.Meta = 2600000;
                        time.MetaAtingimento = ((time.Faturado / time.Meta) * 100);

                        time.Total = time.Salario + time.Comissao;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }

                        Usuarios.Add(time);
                    });

                    #endregion
                }
            }



            return Page();
        }

        public IActionResult OnPost(string Mes,string Ano,string Empresa,string Linha)
        {
            #region Dados
            string Tempo = $"{Mes}/01/{Ano}";

            var date = DateTime.Parse(Tempo);

            //MesAno = date.ToString("MM").ToUpper();
            //this.Ano = date.ToString("yyyy").ToUpper();

            string data = date.ToString("yyyy/MM").Replace("/", "");
            string DataInicio = data + "01";
            string DataFim = data + "31";
            int[] CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };

            var dataini = Convert.ToInt32(DataInicio);
            var cf = new[] { "5551", "6551", "6107", "6109" };

            this.Mes = Mes;
            this.Ano = Ano;

            #endregion

            if (Empresa != "")
            {
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
                        time.Garantia = time.User.Garantia;

                        time.Total = time.Salario + time.Garantia + time.Comissao;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }

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

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }

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

                #region BaixaSub

                var BaixaSubInter = (from SE50 in ProtheusInter.Se5010s
                                     join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                     equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                     join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                     join SC50 in ProtheusInter.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                     from c in sr.DefaultIfEmpty()
                                     where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                     && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                     && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                     && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                     && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                     && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                     && SA10.A1Clinter == "S"
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
                                         Vendedor = c.C5Nomvend,
                                         TipoCliente = SA10.A1Clinter,
                                         Empresa = "INTERMEDIC"
                                     }
                             ).ToList();
                #endregion

                #region BaixaLicitacoes

                var BaixaLicitacoesInter = (from SE50 in ProtheusInter.Se5010s
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
                                            && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                DataPedido = SC50.C5Emissao
                                            }).ToList();

                #endregion

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

                #region BaixaSub

                var BaixaSubDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                     join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                     equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                     join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                     join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                     from c in sr.DefaultIfEmpty()
                                     where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                     && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                     && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                     && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                     && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                     && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                     && SA10.A1Clinter == "S"
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
                                         Vendedor = c.C5Nomvend,
                                         TipoCliente = SA10.A1Clinter,
                                         Empresa = "DENUO"
                                     }
                             ).ToList();

                #endregion

                #region BaixaLicitacoes

                var BaixaLicitacoesDenuo = (from SE50 in ProtheusDenuo.Se5010s
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
                                            && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                DataPedido = SC50.C5Emissao,
                                                Empresa = "DENUO"
                                            }).ToList();

                #endregion

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

                        if (dataini <=  20231231)
                        {

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
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                        i++;
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                    }
                                }
                                else
                                {
                                    time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                    time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                }
                            });
                        }
                        else
                        {
                            time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                            time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                            time.Faturado += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                            time.Faturado += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                            if (usuario == "MICHEL.SAMPAIO")
                            {
                                time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                                time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                            }
                            else
                            {
                                time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                            }
                            int i = 0;
                            Produtos.ForEach(prod =>
                            {
                                if (usuario == "TIAGO.FONSECA")
                                {
                                    if (i == 0)
                                    {
                                        time.FaturadoEquipe = 0;
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        i++;
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                    }
                                }
                                else
                                {
                                    time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                    time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                }
                            });

                        }



                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                        time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                        time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);

                        time.Salario = time.User.Salario;
                        time.Meta = time.User.Meta / 12;
                        time.MetaAtingimento = (((time.Faturado + time.FaturadoProduto) / (time.User.Meta / 12)) * 100);

                        time.Teto = time.User.Teto;
                        time.Garantia = time.User.Garantia;

                        time.Total = time.Salario + time.Garantia + time.Comissao + time.ComissaoEquipe + time.ComissaoProduto;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }


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


                        time.Faturado = BaixaLicitacoesDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaLicitacoesInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                        time.Salario = time.User.Salario;
                        time.Meta = time.User.Meta / 12;
                        time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                        time.Teto = time.User.Teto;

                        time.Garantia = time.User.Garantia;

                        time.Total = time.Salario + time.Garantia + time.Comissao;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }

                        UsuariosLicitacoes.Add(time);
                    }
                    else
                    {

                        time.Faturado = BaixaSubDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaSubInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                        time.Salario = time.User.Salario;
                        time.Meta = time.User.Meta / 12;
                        time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                        time.Teto = time.User.Teto;

                        time.Garantia = time.User.Garantia;

                        time.Total = time.Salario + time.Garantia + time.Comissao;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }

                        UsuariosSub.Add(time);
                    }
                });

                #endregion

                #region ADM

                #region INTERMEDIC

                #region Faturado
                var query5 = (from SD20 in ProtheusInter.Sd2010s
                              join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                              join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                              join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                              join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                              join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                              join SC60 in ProtheusInter.Sc6010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Pedido = SC60.C6Num, Cliente = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                              where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SC60.DELET != "*"
                              && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                              ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                              && SD20.D2Quant != 0 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
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

                var resultadoInter2 = query5.GroupBy(x => new
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

                var DevolucaoInter2 = new List<RelatorioDevolucaoFat>();

                var teste4 = (from SD10 in ProtheusInter.Sd1010s
                              join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                              join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                              join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                              join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                              join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                              where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                               && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                               && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                               && (int)(object)SD10.D1Dtdigit >= 20200701
                               && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K")
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

                DevolucaoInter2 = teste4.Select(x => new RelatorioDevolucaoFat
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

                #endregion

                #region DENUO

                #region Faturado

                var query4 = (from SD20 in ProtheusDenuo.Sd2010s
                              join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                              join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                              join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                              join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                              join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                              where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                 && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                 (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                 && SD20.D2Quant != 0 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SC50.C5Xtipopv != "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
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

                var resultadoDenuo2 = query4.GroupBy(x => new
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
                var DevolucaoDenuo2 = new List<RelatorioDevolucaoFat>();

                var teste3 = (from SD10 in ProtheusDenuo.Sd1010s
                              join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                              join SA30 in ProtheusDenuo.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                              join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                              join SA10 in ProtheusDenuo.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                              join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                              where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
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

                DevolucaoDenuo2 = teste3.Select(x => new RelatorioDevolucaoFat
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

                #endregion

                Users = SGID.TimeADMs.OrderBy(x => x.Integrante).ToList();

                Users.ForEach(x =>
                {
                    var usuario = x.Integrante.ToUpper();


                    var time = new TimeADMRH
                    {
                        User = x,
                    };

                    time.Faturado += resultadoInter2.Sum(x => x.Total) - DevolucaoInter2.Sum(x => x.Total);
                    time.Faturado += resultadoDenuo2.Sum(x => x.Total) - DevolucaoDenuo2.Sum(x => x.Total);

                    time.Salario = time.User.Salario;
                    time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                    time.Meta = 5200000;
                    time.MetaAtingimento = ((time.Faturado / time.Meta) * 100);

                    time.Teto = time.User.Teto;

                    time.Total = time.Salario + time.Comissao;

                    if (time.Teto > 0 && time.Total > time.Teto)
                    {
                        time.Paga = time.Teto;
                    }
                    else
                    {
                        time.Paga = time.Total;
                    }

                    Usuarios.Add(time);
                });

                #endregion
            }
            else
            {
                if (Empresa == "INTERMEDIC")
                {
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

                    #region BaixaSub

                    var BaixaSubInter = (from SE50 in ProtheusInter.Se5010s
                                         join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                         equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                         join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                         join SC50 in ProtheusInter.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                         from c in sr.DefaultIfEmpty()
                                         where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                         && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                         && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                         && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                         && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                         && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                         && SA10.A1Clinter == "S"
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
                                             Vendedor = c.C5Nomvend,
                                             TipoCliente = SA10.A1Clinter,
                                             Empresa = "INTERMEDIC"
                                         }
                                 ).ToList();
                    #endregion

                    #region BaixaLicitacoes

                    var BaixaLicitacoesInter = (from SE50 in ProtheusInter.Se5010s
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
                                                && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                    DataPedido = SC50.C5Emissao
                                                }).ToList();

                    #endregion

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
                            time.Linha = resultadoInter.FirstOrDefault(x => x.Login == usuario)?.Linha;
                            if (time.User.Integrante.ToUpper() == "EDUARDO.ARONI")
                            {
                                time.Linha = "ORTOPEDIA";
                            }
                            else if (time.User.Integrante.ToUpper() == "TIAGO.FONSECA" || time.User.Integrante.ToUpper() == "ARTEMIO.COSTA")
                            {
                                time.Linha = "BUCOMAXILO";
                            }

                            if (dataini <=  20231231)
                            {

                                time.Faturado += resultadoInter.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                                }
                                else
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total);
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
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                    }
                                });
                            }
                            else
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.Faturado += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                                }
                                else
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                }
                                int i = 0;
                                Produtos.ForEach(prod =>
                                {
                                    if (usuario == "TIAGO.FONSECA")
                                    {
                                        if (i == 0)
                                        {
                                            time.FaturadoEquipe = 0;
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                    }
                                });

                            }
                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                            time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = (((time.Faturado + time.FaturadoProduto) / (time.User.Meta / 12)) * 100);
                            time.Garantia = time.User.Garantia;

                            time.Teto = time.User.Teto;

                            time.Total = time.Salario + time.Garantia + time.Comissao + time.ComissaoEquipe + time.ComissaoProduto;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

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

                            time.Faturado = BaixaLicitacoesInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            time.Salario = time.User.Salario;

                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                            time.Teto = time.User.Teto;
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

                            UsuariosLicitacoes.Add(time);
                        }
                        else
                        {

                            time.Faturado = BaixaSubInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                            time.Teto = time.User.Teto;
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

                            UsuariosSub.Add(time);
                        }
                    });

                    #endregion

                    #region ADM

                    #region INTERMEDIC

                    #region Faturado
                    var query5 = (from SD20 in ProtheusInter.Sd2010s
                                  join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                  join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                  join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                  join SC60 in ProtheusInter.Sc6010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Pedido = SC60.C6Num, Cliente = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                                  where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SC60.DELET != "*"
                                  && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                                  ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                  && SD20.D2Quant != 0 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
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

                    var resultadoInter2 = query5.GroupBy(x => new
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

                    var DevolucaoInter2 = new List<RelatorioDevolucaoFat>();

                    var teste4 = (from SD10 in ProtheusInter.Sd1010s
                                  join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                  join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                  join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                  where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                   && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                   && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                   && (int)(object)SD10.D1Dtdigit >= 20200701
                                   && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K")
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

                    DevolucaoInter2 = teste4.Select(x => new RelatorioDevolucaoFat
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

                    #endregion

                    Users = SGID.TimeADMs.OrderBy(x => x.Integrante).ToList();

                    Users.ForEach(x =>
                    {
                        var usuario = x.Integrante.ToUpper();


                        var time = new TimeADMRH
                        {
                            User = x,
                        };

                        time.Faturado += resultadoInter2.Sum(x => x.Total) - DevolucaoInter2.Sum(x => x.Total);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                        time.Salario = time.User.Salario;

                        time.Meta = 2600000;
                        time.MetaAtingimento = ((time.Faturado / time.Meta) * 100);

                        time.Teto = time.User.Teto;

                        time.Total = time.Salario + time.Comissao;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }
                        Usuarios.Add(time);
                    });

                    #endregion
                }
                else
                {
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
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;
                            UsuariosDental.Add(time);
                        }
                        else
                        {
                            time.FaturadoEquipe += Faturamento.Sum(x => x.D2_TOTAL) + Devolucao.Sum(x => x.Total);

                            time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);
                            time.Salario = time.User.Salario;
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

                            UsuariosDental.Add(time);
                        }
                    });

                    #endregion

                    #region Comercial

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

                    #region BaixaSub

                    var BaixaSubDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                         join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                         equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                         join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                         join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                         from c in sr.DefaultIfEmpty()
                                         where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                         && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                         && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                         && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                         && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                         && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                         && SA10.A1Clinter == "S"
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
                                             Vendedor = c.C5Nomvend,
                                             TipoCliente = SA10.A1Clinter,
                                             Empresa = "DENUO"
                                         }
                                 ).ToList();

                    #endregion

                    #region BaixaLicitacoes

                    var BaixaLicitacoesDenuo = (from SE50 in ProtheusDenuo.Se5010s
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
                                                && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                    DataPedido = SC50.C5Emissao,
                                                    Empresa = "DENUO"
                                                }).ToList();

                    #endregion

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

                            if (dataini <=  20231231)
                            {

                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);
                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total);
                                }
                                else
                                {
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
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                    }
                                });
                            }
                            else
                            {
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.Faturado += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                                }
                                else
                                {
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                }
                                int i = 0;
                                Produtos.ForEach(prod =>
                                {
                                    if (usuario == "TIAGO.FONSECA")
                                    {
                                        if (i == 0)
                                        {
                                            time.FaturadoEquipe = 0;
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                    }
                                });

                            }
                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                            time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = (((time.Faturado + time.FaturadoProduto) / (time.User.Meta / 12)) * 100);
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Comissao + time.ComissaoEquipe + time.ComissaoProduto;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

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


                            time.Faturado = BaixaLicitacoesDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

                            UsuariosLicitacoes.Add(time);
                        }
                        else
                        {

                            time.Faturado = BaixaSubDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

                            UsuariosSub.Add(time);
                        }
                    });

                    #endregion

                    #region ADM

                    #region DENUO

                    #region Faturado

                    var query4 = (from SD20 in ProtheusDenuo.Sd2010s
                                  join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                  join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                  join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                  where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                     && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                     (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                     && SD20.D2Quant != 0 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SC50.C5Xtipopv != "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
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

                    var resultadoDenuo2 = query4.GroupBy(x => new
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
                    var DevolucaoDenuo2 = new List<RelatorioDevolucaoFat>();

                    var teste3 = (from SD10 in ProtheusDenuo.Sd1010s
                                  join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SA30 in ProtheusDenuo.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                  join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                  join SA10 in ProtheusDenuo.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                  where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
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

                    DevolucaoDenuo2 = teste3.Select(x => new RelatorioDevolucaoFat
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

                    #endregion

                    Users = SGID.TimeADMs.OrderBy(x => x.Integrante).ToList();

                    Users.ForEach(x =>
                    {
                        var usuario = x.Integrante.ToUpper();


                        var time = new TimeADMRH
                        {
                            User = x,
                        };

                        time.Faturado += resultadoDenuo2.Sum(x => x.Total) - DevolucaoDenuo2.Sum(x => x.Total);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                        time.Salario = time.User.Salario;

                        time.Meta = 2600000;
                        time.MetaAtingimento = ((time.Faturado / time.Meta) * 100);

                        time.Total = time.Salario + time.Comissao;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }

                        Usuarios.Add(time);
                    });

                    #endregion
                }
            }

            switch (Linha)
            {
                case "ORTOPEDIA": 
                    {
                        UsuariosComercial = new List<TimeRH>();
                        ComercialBuco = new List<TimeRH>();
                        ComercialTorax = new List<TimeRH>();
                        ComercialInterior = new List<TimeRH>();
                        UsuariosLicitacoes = new List<TimeRH>();
                        UsuariosSub = new List<TimeRH>();
                        Usuarios = new List<TimeADMRH>();
                        UsuariosDental = new List<TimeDentalRH>();
                        break;
                    }
                case "INTERIOR":
                    {
                        UsuariosComercial = new List<TimeRH>();
                        ComercialOrtopedia = new List<TimeRH>();
                        ComercialBuco = new List<TimeRH>();
                        ComercialTorax = new List<TimeRH>();
                        UsuariosLicitacoes = new List<TimeRH>();
                        UsuariosSub = new List<TimeRH>();
                        Usuarios = new List<TimeADMRH>();
                        UsuariosDental = new List<TimeDentalRH>();
                        break;
                    }
                case "LICITACOES":
                    {
                        UsuariosComercial = new List<TimeRH>();
                        ComercialOrtopedia = new List<TimeRH>();
                        ComercialBuco = new List<TimeRH>();
                        ComercialTorax = new List<TimeRH>();
                        ComercialInterior = new List<TimeRH>();
                        UsuariosSub = new List<TimeRH>();
                        Usuarios = new List<TimeADMRH>();
                        UsuariosDental = new List<TimeDentalRH>();
                        break;
                    }
                case "SUB":
                    {
                        UsuariosComercial = new List<TimeRH>();
                        ComercialOrtopedia = new List<TimeRH>();
                        ComercialBuco = new List<TimeRH>();
                        ComercialTorax = new List<TimeRH>();
                        ComercialInterior = new List<TimeRH>();
                        UsuariosLicitacoes = new List<TimeRH>();
                        Usuarios = new List<TimeADMRH>();
                        UsuariosDental = new List<TimeDentalRH>();
                        break;
                    }
                case "BUCO":
                    {
                        UsuariosComercial = new List<TimeRH>();
                        ComercialOrtopedia = new List<TimeRH>();
                        ComercialTorax = new List<TimeRH>();
                        ComercialInterior = new List<TimeRH>();
                        UsuariosLicitacoes = new List<TimeRH>();
                        UsuariosSub = new List<TimeRH>();
                        Usuarios = new List<TimeADMRH>();
                        UsuariosDental = new List<TimeDentalRH>();
                        break;
                    }
                case "TORAX":
                    {
                        UsuariosComercial = new List<TimeRH>();
                        ComercialOrtopedia = new List<TimeRH>();
                        ComercialBuco = new List<TimeRH>();
                        ComercialInterior = new List<TimeRH>();
                        UsuariosLicitacoes = new List<TimeRH>();
                        UsuariosSub = new List<TimeRH>();
                        Usuarios = new List<TimeADMRH>();
                        UsuariosDental = new List<TimeDentalRH>();
                        break;
                    }
                case "DENTAL":
                    {
                        UsuariosComercial = new List<TimeRH>();
                        ComercialOrtopedia = new List<TimeRH>();
                        ComercialBuco = new List<TimeRH>();
                        ComercialTorax = new List<TimeRH>();
                        ComercialInterior = new List<TimeRH>();
                        UsuariosLicitacoes = new List<TimeRH>();
                        UsuariosSub = new List<TimeRH>();
                        Usuarios = new List<TimeADMRH>();
                        break;
                    }
                case "ADMINISTRATIVO":
                    {
                        UsuariosComercial = new List<TimeRH>();
                        ComercialOrtopedia = new List<TimeRH>();
                        ComercialBuco = new List<TimeRH>();
                        ComercialTorax = new List<TimeRH>();
                        ComercialInterior = new List<TimeRH>();
                        UsuariosLicitacoes = new List<TimeRH>();
                        UsuariosSub = new List<TimeRH>();
                        UsuariosDental = new List<TimeDentalRH>();
                        break;
                    }
            }

            return Page();
        }

        public IActionResult OnPostExport(string Mes, string Ano, string Empresa, string Linha)
        {
            try
            {
                #region Consulta
                #region Dados
                string Tempo = $"{Mes}/01/{Ano}";

                var date = DateTime.Parse(Tempo);

                //MesAno = date.ToString("MM").ToUpper();
                //this.Ano = date.ToString("yyyy").ToUpper();

                string data = date.ToString("yyyy/MM").Replace("/", "");
                string DataInicio = data + "01";
                string DataFim = data + "31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };

                var dataini = Convert.ToInt32(DataInicio);
                var cf = new[] { "5551", "6551", "6107", "6109" };

                this.Mes = Mes;
                this.Ano = Ano;

                #endregion

                if (Empresa != "")
                {
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
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

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

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

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

                    #region BaixaSub

                    var BaixaSubInter = (from SE50 in ProtheusInter.Se5010s
                                         join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                         equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                         join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                         join SC50 in ProtheusInter.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                         from c in sr.DefaultIfEmpty()
                                         where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                         && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                         && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                         && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                         && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                         && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                         && SA10.A1Clinter == "S"
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
                                             Vendedor = c.C5Nomvend,
                                             TipoCliente = SA10.A1Clinter,
                                             Empresa = "INTERMEDIC"
                                         }
                                 ).ToList();
                    #endregion

                    #region BaixaLicitacoes

                    var BaixaLicitacoesInter = (from SE50 in ProtheusInter.Se5010s
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
                                                && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                    DataPedido = SC50.C5Emissao
                                                }).ToList();

                    #endregion

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

                    #region BaixaSub

                    var BaixaSubDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                         join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                         equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                         join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                         join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                         from c in sr.DefaultIfEmpty()
                                         where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                         && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                         && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                         && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                         && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                         && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                         && SA10.A1Clinter == "S"
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
                                             Vendedor = c.C5Nomvend,
                                             TipoCliente = SA10.A1Clinter,
                                             Empresa = "DENUO"
                                         }
                                 ).ToList();

                    #endregion

                    #region BaixaLicitacoes

                    var BaixaLicitacoesDenuo = (from SE50 in ProtheusDenuo.Se5010s
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
                                                && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                    DataPedido = SC50.C5Emissao,
                                                    Empresa = "DENUO"
                                                }).ToList();

                    #endregion

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

                            if (dataini <=  20231231)
                            {

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
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                    }
                                });
                            }
                            else
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                time.Faturado += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                time.Faturado += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                                }
                                else
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                }
                                int i = 0;
                                Produtos.ForEach(prod =>
                                {
                                    if (usuario == "TIAGO.FONSECA")
                                    {
                                        if (i == 0)
                                        {
                                            time.FaturadoEquipe = 0;
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                    }
                                });

                            }



                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                            time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = (((time.Faturado + time.FaturadoProduto) / (time.User.Meta / 12)) * 100);

                            time.Teto = time.User.Teto;
                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao + time.ComissaoEquipe + time.ComissaoProduto;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }


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


                            time.Faturado = BaixaLicitacoesDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaLicitacoesInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                            time.Teto = time.User.Teto;

                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

                            UsuariosLicitacoes.Add(time);
                        }
                        else
                        {

                            time.Faturado = BaixaSubDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaSubInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            time.Salario = time.User.Salario;
                            time.Meta = time.User.Meta / 12;
                            time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                            time.Teto = time.User.Teto;

                            time.Garantia = time.User.Garantia;

                            time.Total = time.Salario + time.Garantia + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

                            UsuariosSub.Add(time);
                        }
                    });

                    #endregion

                    #region ADM


                    #region INTERMEDIC

                    #region Faturado
                    var query5 = (from SD20 in ProtheusInter.Sd2010s
                                 join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                 join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                 join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                 join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                 join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                 join SC60 in ProtheusInter.Sc6010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Pedido = SC60.C6Num, Cliente = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                                 where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SC60.DELET != "*"
                                 && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                                 ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                 && SD20.D2Quant != 0 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
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

                    var resultadoInter2 = query5.GroupBy(x => new
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

                    var DevolucaoInter2 = new List<RelatorioDevolucaoFat>();

                    var teste4 = (from SD10 in ProtheusInter.Sd1010s
                                 join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                 join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                 join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                 join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                 join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                 where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                  && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                  && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                  && (int)(object)SD10.D1Dtdigit >= 20200701
                                  && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K")
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

                    DevolucaoInter2 = teste4.Select(x => new RelatorioDevolucaoFat
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

                    #endregion

                    #region DENUO

                    #region Faturado

                    var query4 = (from SD20 in ProtheusDenuo.Sd2010s
                                  join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                  join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                  join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                  where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                     && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                     (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                     && SD20.D2Quant != 0 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SC50.C5Xtipopv != "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
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

                    var resultadoDenuo2 = query4.GroupBy(x => new
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
                    var DevolucaoDenuo2 = new List<RelatorioDevolucaoFat>();

                    var teste3 = (from SD10 in ProtheusDenuo.Sd1010s
                                  join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                  join SA30 in ProtheusDenuo.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                  join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                  join SA10 in ProtheusDenuo.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                  where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
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

                    DevolucaoDenuo2 = teste3.Select(x => new RelatorioDevolucaoFat
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

                    #endregion



                    Users = SGID.TimeADMs.OrderBy(x => x.Integrante).ToList();

                    Users.ForEach(x =>
                    {
                        var usuario = x.Integrante.ToUpper();


                        var time = new TimeADMRH
                        {
                            User = x,
                        };

                        time.Faturado += resultadoInter2.Sum(x => x.Total) - DevolucaoInter2.Sum(x => x.Total);
                        time.Faturado += resultadoDenuo2.Sum(x => x.Total) - DevolucaoDenuo2.Sum(x => x.Total);

                        time.Salario = time.User.Salario;
                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                        time.Meta = 5200000;
                        time.MetaAtingimento = ((time.Faturado / time.Meta) * 100);

                        time.Teto = time.User.Teto;

                        time.Total = time.Salario + time.Comissao;

                        if (time.Teto > 0 && time.Total > time.Teto)
                        {
                            time.Paga = time.Teto;
                        }
                        else
                        {
                            time.Paga = time.Total;
                        }

                        Usuarios.Add(time);
                    });

                    #endregion
                }
                else
                {
                    if (Empresa == "INTERMEDIC")
                    {
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

                        #region BaixaSub

                        var BaixaSubInter = (from SE50 in ProtheusInter.Se5010s
                                             join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                             equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                             join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                             join SC50 in ProtheusInter.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                             from c in sr.DefaultIfEmpty()
                                             where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                             && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                             && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                             && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                             && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                             && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                             && SA10.A1Clinter == "S"
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
                                                 Vendedor = c.C5Nomvend,
                                                 TipoCliente = SA10.A1Clinter,
                                                 Empresa = "INTERMEDIC"
                                             }
                                     ).ToList();
                        #endregion

                        #region BaixaLicitacoes

                        var BaixaLicitacoesInter = (from SE50 in ProtheusInter.Se5010s
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
                                                    && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                        DataPedido = SC50.C5Emissao
                                                    }).ToList();

                        #endregion

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
                                time.Linha = resultadoInter.FirstOrDefault(x => x.Login == usuario)?.Linha;
                                if (time.User.Integrante.ToUpper() == "EDUARDO.ARONI")
                                {
                                    time.Linha = "ORTOPEDIA";
                                }
                                else if (time.User.Integrante.ToUpper() == "TIAGO.FONSECA" || time.User.Integrante.ToUpper() == "ARTEMIO.COSTA")
                                {
                                    time.Linha = "BUCOMAXILO";
                                }

                                if (dataini <=  20231231)
                                {

                                    time.Faturado += resultadoInter.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                                    if (usuario == "MICHEL.SAMPAIO")
                                    {
                                        time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                                    }
                                    else
                                    {
                                        time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total);
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
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        }
                                    });
                                }
                                else
                                {
                                    time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.Faturado += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                    if (usuario == "MICHEL.SAMPAIO")
                                    {
                                        time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                                    }
                                    else
                                    {
                                        time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                    }
                                    int i = 0;
                                    Produtos.ForEach(prod =>
                                    {
                                        if (usuario == "TIAGO.FONSECA")
                                        {
                                            if (i == 0)
                                            {
                                                time.FaturadoEquipe = 0;
                                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        }
                                    });

                                }
                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);

                                time.Salario = time.User.Salario;
                                time.Meta = time.User.Meta / 12;
                                time.MetaAtingimento = (((time.Faturado + time.FaturadoProduto) / (time.User.Meta / 12)) * 100);
                                time.Garantia = time.User.Garantia;

                                time.Teto = time.User.Teto;

                                time.Total = time.Salario + time.Garantia + time.Comissao + time.ComissaoEquipe + time.ComissaoProduto;

                                if (time.Teto > 0 && time.Total > time.Teto)
                                {
                                    time.Paga = time.Teto;
                                }
                                else
                                {
                                    time.Paga = time.Total;
                                }

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

                                time.Faturado = BaixaLicitacoesInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                                time.Salario = time.User.Salario;

                                time.Meta = time.User.Meta / 12;
                                time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                                time.Teto = time.User.Teto;
                                time.Garantia = time.User.Garantia;

                                time.Total = time.Salario + time.Garantia + time.Comissao;

                                if (time.Teto > 0 && time.Total > time.Teto)
                                {
                                    time.Paga = time.Teto;
                                }
                                else
                                {
                                    time.Paga = time.Total;
                                }

                                UsuariosLicitacoes.Add(time);
                            }
                            else
                            {

                                time.Faturado = BaixaSubInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                                time.Salario = time.User.Salario;
                                time.Meta = time.User.Meta / 12;
                                time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);

                                time.Teto = time.User.Teto;
                                time.Garantia = time.User.Garantia;

                                time.Total = time.Salario + time.Comissao;

                                if (time.Teto > 0 && time.Total > time.Teto)
                                {
                                    time.Paga = time.Teto;
                                }
                                else
                                {
                                    time.Paga = time.Total;
                                }

                                UsuariosSub.Add(time);
                            }
                        });

                        #endregion

                        #region ADM

                        #region INTERMEDIC

                        #region Faturado
                        var query5 = (from SD20 in ProtheusInter.Sd2010s
                                      join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                      join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                      join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                      join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                      join SC60 in ProtheusInter.Sc6010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Pedido = SC60.C6Num, Cliente = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                                      where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SC60.DELET != "*"
                                      && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                                      ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                      && SD20.D2Quant != 0 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
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

                        var resultadoInter2 = query5.GroupBy(x => new
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

                        var DevolucaoInter2 = new List<RelatorioDevolucaoFat>();

                        var teste4 = (from SD10 in ProtheusInter.Sd1010s
                                      join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                      join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                      join SC50 in ProtheusInter.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                      join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                      where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                                       && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                       && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                                       && (int)(object)SD10.D1Dtdigit >= 20200701
                                       && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K")
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

                        DevolucaoInter2 = teste4.Select(x => new RelatorioDevolucaoFat
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

                        #endregion

                        Users = SGID.TimeADMs.OrderBy(x => x.Integrante).ToList();

                        Users.ForEach(x =>
                        {
                            var usuario = x.Integrante.ToUpper();


                            var time = new TimeADMRH
                            {
                                User = x,
                            };

                            time.Faturado += resultadoInter2.Sum(x => x.Total) - DevolucaoInter2.Sum(x => x.Total);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            time.Salario = time.User.Salario;

                            time.Meta = 2600000;
                            time.MetaAtingimento = ((time.Faturado / time.Meta) * 100);

                            time.Teto = time.User.Teto;

                            time.Total = time.Salario + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }
                            Usuarios.Add(time);
                        });

                        #endregion
                    }
                    else
                    {
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
                                time.Garantia = time.User.Garantia;

                                time.Total = time.Salario + time.Garantia + time.Comissao;
                                UsuariosDental.Add(time);
                            }
                            else
                            {
                                time.FaturadoEquipe += Faturamento.Sum(x => x.D2_TOTAL) + Devolucao.Sum(x => x.Total);

                                time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                time.Meta = time.User.Meta / 12;
                                time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);
                                time.Salario = time.User.Salario;
                                time.Garantia = time.User.Garantia;

                                time.Total = time.Salario + time.Garantia + time.Comissao;

                                if (time.Teto > 0 && time.Total > time.Teto)
                                {
                                    time.Paga = time.Teto;
                                }
                                else
                                {
                                    time.Paga = time.Total;
                                }

                                UsuariosDental.Add(time);
                            }
                        });

                        #endregion

                        #region Comercial

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

                        #region BaixaSub

                        var BaixaSubDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                             join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                             equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                             join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                             join SC50 in ProtheusDenuo.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                                             from c in sr.DefaultIfEmpty()
                                             where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                                             && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                             && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                             && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                             && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                             && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                             && SA10.A1Clinter == "S"
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
                                                 Vendedor = c.C5Nomvend,
                                                 TipoCliente = SA10.A1Clinter,
                                                 Empresa = "DENUO"
                                             }
                                     ).ToList();

                        #endregion

                        #region BaixaLicitacoes

                        var BaixaLicitacoesDenuo = (from SE50 in ProtheusDenuo.Se5010s
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
                                                    && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
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
                                                        DataPedido = SC50.C5Emissao,
                                                        Empresa = "DENUO"
                                                    }).ToList();

                        #endregion

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

                                if (dataini <=  20231231)
                                {

                                    time.Faturado += resultadoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);
                                    if (usuario == "MICHEL.SAMPAIO")
                                    {
                                        time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total);
                                    }
                                    else
                                    {
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
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim()).Sum(x => x.Total);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        }
                                    });
                                }
                                else
                                {
                                    time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                    time.Faturado += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                    if (usuario == "MICHEL.SAMPAIO")
                                    {
                                        time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == "ANDRE.SALES").Sum(x => x.TotalBaixado);
                                    }
                                    else
                                    {
                                        time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                    }
                                    int i = 0;
                                    Produtos.ForEach(prod =>
                                    {
                                        if (usuario == "TIAGO.FONSECA")
                                        {
                                            if (i == 0)
                                            {
                                                time.FaturadoEquipe = 0;
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20221231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha == prod.Produto.Trim()).Sum(x => x.TotalBaixado);
                                        }
                                    });

                                }
                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);

                                time.Salario = time.User.Salario;
                                time.Meta = time.User.Meta / 12;
                                time.MetaAtingimento = (((time.Faturado + time.FaturadoProduto) / (time.User.Meta / 12)) * 100);
                                time.Garantia = time.User.Garantia;

                                time.Total = time.Salario + time.Comissao + time.ComissaoEquipe + time.ComissaoProduto;

                                if (time.Teto > 0 && time.Total > time.Teto)
                                {
                                    time.Paga = time.Teto;
                                }
                                else
                                {
                                    time.Paga = time.Total;
                                }

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


                                time.Faturado = BaixaLicitacoesDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                                time.Salario = time.User.Salario;
                                time.Meta = time.User.Meta / 12;
                                time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);
                                time.Garantia = time.User.Garantia;

                                time.Total = time.Salario + time.Garantia + time.Comissao;

                                if (time.Teto > 0 && time.Total > time.Teto)
                                {
                                    time.Paga = time.Teto;
                                }
                                else
                                {
                                    time.Paga = time.Total;
                                }

                                UsuariosLicitacoes.Add(time);
                            }
                            else
                            {

                                time.Faturado = BaixaSubDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                                time.Salario = time.User.Salario;
                                time.Meta = time.User.Meta / 12;
                                time.MetaAtingimento = ((time.Faturado / (time.User.Meta / 12)) * 100);
                                time.Garantia = time.User.Garantia;

                                time.Total = time.Salario + time.Garantia + time.Comissao;

                                if (time.Teto > 0 && time.Total > time.Teto)
                                {
                                    time.Paga = time.Teto;
                                }
                                else
                                {
                                    time.Paga = time.Total;
                                }

                                UsuariosSub.Add(time);
                            }
                        });

                        #endregion

                        #region ADM

                        #region DENUO

                        #region Faturado

                        var query4 = (from SD20 in ProtheusDenuo.Sd2010s
                                      join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                      join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                      join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                      join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                      where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                         && SD20.D2Quant != 0 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SC50.C5Xtipopv != "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
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

                        var resultadoDenuo2 = query4.GroupBy(x => new
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
                        var DevolucaoDenuo2 = new List<RelatorioDevolucaoFat>();

                        var teste3 = (from SD10 in ProtheusDenuo.Sd1010s
                                      join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                      join SA30 in ProtheusDenuo.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                      join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                      join SA10 in ProtheusDenuo.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                      where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
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

                        DevolucaoDenuo2 = teste3.Select(x => new RelatorioDevolucaoFat
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

                        #endregion

                        Users = SGID.TimeADMs.OrderBy(x => x.Integrante).ToList();

                        Users.ForEach(x =>
                        {
                            var usuario = x.Integrante.ToUpper();


                            var time = new TimeADMRH
                            {
                                User = x,
                            };

                            time.Faturado += resultadoDenuo.Sum(x => x.Total) - DevolucaoDenuo.Sum(x => x.Total);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            time.Salario = time.User.Salario;

                            time.Meta = 2600000;
                            time.MetaAtingimento = ((time.Faturado / time.Meta) * 100);

                            time.Total = time.Salario + time.Comissao;

                            if (time.Teto > 0 && time.Total > time.Teto)
                            {
                                time.Paga = time.Teto;
                            }
                            else
                            {
                                time.Paga = time.Total;
                            }

                            Usuarios.Add(time);
                        });

                        #endregion
                    }
                }

                switch (Linha)
                {
                    case "ORTOPEDIA":
                        {
                            UsuariosComercial = new List<TimeRH>();
                            ComercialBuco = new List<TimeRH>();
                            ComercialTorax = new List<TimeRH>();
                            ComercialInterior = new List<TimeRH>();
                            UsuariosLicitacoes = new List<TimeRH>();
                            UsuariosSub = new List<TimeRH>();
                            Usuarios = new List<TimeADMRH>();
                            UsuariosDental = new List<TimeDentalRH>();
                            break;
                        }
                    case "INTERIOR":
                        {
                            UsuariosComercial = new List<TimeRH>();
                            ComercialOrtopedia = new List<TimeRH>();
                            ComercialBuco = new List<TimeRH>();
                            ComercialTorax = new List<TimeRH>();
                            UsuariosLicitacoes = new List<TimeRH>();
                            UsuariosSub = new List<TimeRH>();
                            Usuarios = new List<TimeADMRH>();
                            UsuariosDental = new List<TimeDentalRH>();
                            break;
                        }
                    case "LICITACOES":
                        {
                            UsuariosComercial = new List<TimeRH>();
                            ComercialOrtopedia = new List<TimeRH>();
                            ComercialBuco = new List<TimeRH>();
                            ComercialTorax = new List<TimeRH>();
                            ComercialInterior = new List<TimeRH>();
                            UsuariosSub = new List<TimeRH>();
                            Usuarios = new List<TimeADMRH>();
                            UsuariosDental = new List<TimeDentalRH>();
                            break;
                        }
                    case "SUB":
                        {
                            UsuariosComercial = new List<TimeRH>();
                            ComercialOrtopedia = new List<TimeRH>();
                            ComercialBuco = new List<TimeRH>();
                            ComercialTorax = new List<TimeRH>();
                            ComercialInterior = new List<TimeRH>();
                            UsuariosLicitacoes = new List<TimeRH>();
                            Usuarios = new List<TimeADMRH>();
                            UsuariosDental = new List<TimeDentalRH>();
                            break;
                        }
                    case "BUCO":
                        {
                            UsuariosComercial = new List<TimeRH>();
                            ComercialOrtopedia = new List<TimeRH>();
                            ComercialTorax = new List<TimeRH>();
                            ComercialInterior = new List<TimeRH>();
                            UsuariosLicitacoes = new List<TimeRH>();
                            UsuariosSub = new List<TimeRH>();
                            Usuarios = new List<TimeADMRH>();
                            UsuariosDental = new List<TimeDentalRH>();
                            break;
                        }
                    case "TORAX":
                        {
                            UsuariosComercial = new List<TimeRH>();
                            ComercialOrtopedia = new List<TimeRH>();
                            ComercialBuco = new List<TimeRH>();
                            ComercialInterior = new List<TimeRH>();
                            UsuariosLicitacoes = new List<TimeRH>();
                            UsuariosSub = new List<TimeRH>();
                            Usuarios = new List<TimeADMRH>();
                            UsuariosDental = new List<TimeDentalRH>();
                            break;
                        }
                    case "DENTAL":
                        {
                            UsuariosComercial = new List<TimeRH>();
                            ComercialOrtopedia = new List<TimeRH>();
                            ComercialBuco = new List<TimeRH>();
                            ComercialTorax = new List<TimeRH>();
                            ComercialInterior = new List<TimeRH>();
                            UsuariosLicitacoes = new List<TimeRH>();
                            UsuariosSub = new List<TimeRH>();
                            Usuarios = new List<TimeADMRH>();
                            break;
                        }
                    case "ADMINISTRATIVO":
                        {
                            UsuariosComercial = new List<TimeRH>();
                            ComercialOrtopedia = new List<TimeRH>();
                            ComercialBuco = new List<TimeRH>();
                            ComercialTorax = new List<TimeRH>();
                            ComercialInterior = new List<TimeRH>();
                            UsuariosLicitacoes = new List<TimeRH>();
                            UsuariosSub = new List<TimeRH>();
                            UsuariosDental = new List<TimeDentalRH>();
                            break;
                        }
                }

                #endregion

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("SubDistribuidor Faturado");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "SubDistribuidor Faturado");

                sheet.Cells[1, 1].Value = "Nome";
                sheet.Cells[1, 2].Value = "Mês/Ano";
                sheet.Cells[1, 3].Value = "Faturamento Mes";
                sheet.Cells[1, 4].Value = "Recebimento Mes";
                sheet.Cells[1, 5].Value = "Meta Mensal";
                sheet.Cells[1, 6].Value = "Atingimento Meta";
                sheet.Cells[1, 7].Value = "Salario Base";
                sheet.Cells[1, 8].Value = "Garantia";
                sheet.Cells[1, 9].Value = "Comissão Vendedor";
                sheet.Cells[1, 10].Value = "Comissão Equipe";
                sheet.Cells[1, 11].Value = "Total Bruto Remuneração";
                sheet.Cells[1, 12].Value = "Teto";
                sheet.Cells[1, 13].Value = "Total Geral";

                int i = 2;

                foreach(var item in UsuariosLicitacoes)
                {

                    sheet.Cells[i, 1].Value = item.User.Integrante.ToUpper().Split("@")[0].Replace(".", " ");
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = 0;
                    sheet.Cells[i, 4].Value = string.Format("{0:N2}", item.Faturado);
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", item.Meta);
                    if(item.Meta > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", item.MetaAtingimento);
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", item.Salario);
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", item.Garantia);
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", item.Comissao);
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", item.ComissaoEquipe + item.ComissaoProduto);
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", item.Total);
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", item.Teto);
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", item.Paga);
                    
                    i++;
                }

                //background-color:#808080;color:white
                if (UsuariosLicitacoes.Count > 0)
                {
                    sheet.Cells[i, 1].Value = "LICITAÇÕES";
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = 0;
                    sheet.Cells[i, 4].Value = string.Format("{0:N2}", UsuariosLicitacoes.Sum(x => x.Faturado));
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", UsuariosLicitacoes.Sum(x => x.Meta));
                    if(UsuariosLicitacoes.Sum(x => x.Meta) > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", ((UsuariosLicitacoes.Sum(x => x.Faturado) / UsuariosLicitacoes.Sum(x => x.Meta)) * 100));
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", UsuariosLicitacoes.Sum(x => x.Salario));
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", UsuariosLicitacoes.Sum(x => x.Garantia));
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", UsuariosLicitacoes.Sum(x => x.Comissao));
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", UsuariosLicitacoes.Sum(x => x.ComissaoEquipe) + UsuariosLicitacoes.Sum(x => x.ComissaoProduto));
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", UsuariosLicitacoes.Sum(x => x.Total));
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", UsuariosLicitacoes.Sum(x => x.Teto));
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", UsuariosLicitacoes.Sum(x => x.Paga));
                    i++;
                }

                foreach(var item in UsuariosSub)
                {

                    sheet.Cells[i, 1].Value = item.User.Integrante.ToUpper().Split("@")[0].Replace(".", " ");
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = 0;
                    sheet.Cells[i, 4].Value = string.Format("{0:N2}", item.Faturado);
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", item.Meta);
                    
                    if(item.Meta > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", item.MetaAtingimento);
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", item.Salario);
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", item.Garantia);
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", item.Comissao);
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", item.ComissaoEquipe + item.ComissaoProduto);
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", item.Total);
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", item.Teto);
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", item.Paga);
                    i++;
                }

                if(UsuariosSub.Count > 0)
                {
                    sheet.Cells[i, 1].Value = "SUB-DISTRIBUIDOR";
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = 0;
                    sheet.Cells[i, 4].Value = string.Format("{0:N2}", UsuariosSub.First().Faturado);
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Meta));
                    if(UsuariosSub.Sum(x => x.Meta) > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", ((UsuariosSub.First().Faturado / UsuariosSub.Sum(x => x.Meta)) * 100));
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }

                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Salario));
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Garantia));
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Comissao));
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.ComissaoEquipe) + UsuariosSub.Sum(x => x.ComissaoProduto));
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Total));
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Teto));
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Paga));
                    i++;
                }

                foreach (var item in ComercialOrtopedia)
                {
                    sheet.Cells[i, 1].Value = item.User.Integrante.ToUpper().Split("@")[0].Replace(".", " ");
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", item.Faturado);
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", item.Meta);
                    if(item.Meta > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", item.MetaAtingimento);
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", item.Salario);
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", item.Garantia);
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", item.Comissao);
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", item.ComissaoEquipe + item.ComissaoProduto);
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", item.Total);
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", item.Teto);
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", item.Paga);
                    i++;
                }

                if(ComercialOrtopedia.Count > 0)
                {
                    sheet.Cells[i, 1].Value = "ORTOPEDIA";
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", ComercialOrtopedia.Sum(x => x.Faturado));
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", ComercialOrtopedia.Sum(x => x.Meta));
                    if(ComercialOrtopedia.Sum(x => x.Meta) > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", ((ComercialOrtopedia.Sum(x => x.Faturado) / ComercialOrtopedia.Sum(x => x.Meta)) * 100));
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", ComercialOrtopedia.Sum(x => x.Salario));
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", ComercialOrtopedia.Sum(x => x.Garantia));
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", ComercialOrtopedia.Sum(x => x.Comissao));
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", ComercialOrtopedia.Sum(x => x.ComissaoEquipe) + ComercialOrtopedia.Sum(x => x.ComissaoProduto));
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", ComercialOrtopedia.Sum(x => x.Total));
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", ComercialOrtopedia.Sum(x => x.Teto));
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", ComercialOrtopedia.Sum(x => x.Paga));
                    i++;
                }

                foreach (var item in ComercialBuco)
                {
                    sheet.Cells[i, 1].Value = item.User.Integrante.ToUpper().Split("@")[0].Replace(".", " ");
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", item.Faturado);
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", item.Meta);
                    if(item.Meta > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", item.MetaAtingimento);
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", item.Salario);
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", item.Garantia);
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", item.Comissao);
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", item.ComissaoEquipe + item.ComissaoProduto);
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", item.Total);
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", item.Teto);
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", item.Paga);
                    i++;
                }

                if (ComercialBuco.Count > 0)
                {
                    sheet.Cells[i, 1].Value = "BUCOMAXILO";
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", ComercialBuco.Sum(x => x.Faturado));
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", ComercialBuco.Sum(x => x.Meta));
                    if(ComercialBuco.Sum(x => x.Meta) > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", ((ComercialBuco.Sum(x => x.Faturado) / ComercialBuco.Sum(x => x.Meta)) * 100));
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", ComercialBuco.Sum(x => x.Salario));
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", ComercialBuco.Sum(x => x.Garantia));
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", ComercialBuco.Sum(x => x.Comissao));
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", ComercialBuco.Sum(x => x.ComissaoEquipe) + ComercialBuco.Sum(x => x.ComissaoProduto));
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", ComercialBuco.Sum(x => x.Total));
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", ComercialBuco.Sum(x => x.Teto));
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", ComercialBuco.Sum(x => x.Paga));
                    i++;
                }

                foreach(var item in ComercialTorax)
                {
                    sheet.Cells[i, 1].Value = item.User.Integrante.ToUpper().Split("@")[0].Replace(".", " ");
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", item.Faturado);
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", item.Meta);
                    if(item.Meta > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", item.MetaAtingimento);
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", item.Salario);
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", item.Garantia);
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", item.Comissao);
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", item.ComissaoEquipe + item.ComissaoProduto);
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", item.Total);
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", item.Teto);
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", item.Paga);
                    i++;
                }

                if(ComercialTorax.Count > 0)
                {
                    sheet.Cells[i, 1].Value = "TORAX";
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", ComercialTorax.Sum(x => x.Faturado));
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", ComercialTorax.Sum(x => x.Meta));
                    if(ComercialTorax.Sum(x => x.Meta) > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", ((ComercialTorax.Sum(x => x.Faturado) / ComercialTorax.Sum(x => x.Meta)) * 100));
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", ComercialTorax.Sum(x => x.Salario));
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", ComercialTorax.Sum(x => x.Garantia));
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", ComercialTorax.Sum(x => x.Comissao));
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", ComercialTorax.Sum(x => x.ComissaoEquipe) + ComercialBuco.Sum(x => x.ComissaoProduto));
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", ComercialTorax.Sum(x => x.Total));
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", ComercialTorax.Sum(x => x.Teto));
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", ComercialTorax.Sum(x => x.Paga));
                    i++;
                }

                foreach(var item in ComercialInterior)
                {
                    sheet.Cells[i, 1].Value = item.User.Integrante.ToUpper().Split("@")[0].Replace(".", " ");
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", item.Faturado);
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", item.Meta);
                    if(item.Meta > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", item.MetaAtingimento);
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", item.Salario);
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", item.Garantia);
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", item.Comissao);
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", item.ComissaoEquipe + item.ComissaoProduto);
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", item.Total);
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", item.Teto);
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", item.Paga);
                    i++;
                }

                if(ComercialInterior.Count > 0)
                {
                    sheet.Cells[i, 1].Value = "INTERIOR";
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", ComercialInterior.Sum(x => x.Faturado));
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", ComercialInterior.Sum(x => x.Meta));
                    if(ComercialInterior.Sum(x => x.Meta) > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", ((ComercialInterior.Sum(x => x.Faturado) / ComercialInterior.Sum(x => x.Meta)) * 100));
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", ComercialInterior.Sum(x => x.Salario));
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", ComercialInterior.Sum(x => x.Garantia));
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", ComercialInterior.Sum(x => x.Comissao));
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", ComercialInterior.Sum(x => x.ComissaoEquipe) + ComercialInterior.Sum(x => x.ComissaoProduto));
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", ComercialInterior.Sum(x => x.Total));
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", ComercialInterior.Sum(x => x.Teto));
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", ComercialInterior.Sum(x => x.Paga));
                    i++;
                }

                foreach(var item in UsuariosDental)
                {
                    sheet.Cells[i, 1].Value = item.User.Integrante.ToUpper().Split("@")[0].Replace(".", " ");
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", item.Faturado);
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", item.Meta);
                    if(item.Meta > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", item.MetaAtingimento);
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", item.Salario);
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", item.Garantia);
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", item.Comissao);
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", item.ComissaoEquipe);
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", item.Total);
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", item.Teto);
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", item.Paga);
                    i++;
                }

                if(UsuariosDental.Count > 0)
                {
                    sheet.Cells[i, 1].Value = "DENTAL";
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", UsuariosDental.Sum(x => x.Faturado));
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", UsuariosDental.Sum(x => x.Meta));
                    if(UsuariosDental.Sum(x => x.Meta) > 0)
                    {
                        sheet.Cells[i, 6].Value = string.Format("{0:N2}", ((UsuariosDental.Sum(x => x.Faturado) / UsuariosDental.Sum(x => x.Meta)) * 100));
                    }
                    else
                    {
                        sheet.Cells[i, 6].Value = 0;
                    }
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", UsuariosDental.Sum(x => x.Salario));
                    sheet.Cells[i, 8].Value = string.Format("{0:N2}", UsuariosDental.Sum(x => x.Garantia));
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", UsuariosDental.Sum(x => x.Comissao));
                    sheet.Cells[i, 10].Value = string.Format("{0:N2}", UsuariosDental.Sum(x => x.ComissaoEquipe));
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", UsuariosDental.Sum(x => x.Total));
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", UsuariosDental.Sum(x => x.Teto));
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", UsuariosDental.Sum(x => x.Paga));
                    i++;
                }

                foreach(var item in Usuarios)
                {
                    sheet.Cells[i, 1].Value = item.User.Integrante.ToUpper().Split("@")[0].Replace(".", " ");
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", item.Faturado);
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", item.Meta);
                    sheet.Cells[i, 6].Value = string.Format("{0:N2}", item.MetaAtingimento);
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", item.Salario);
                    sheet.Cells[i, 8].Value = 0;
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", item.Comissao);
                    sheet.Cells[i, 10].Value = 0;
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", item.Total);
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", item.Teto);
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", item.Paga);
                    i++;
                }

                if(Usuarios.Count > 0)
                {
                    sheet.Cells[i, 1].Value = "ADMINISTRATIVO";
                    sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                    sheet.Cells[i, 3].Value = string.Format("{0:N2}", Usuarios.First().Faturado);
                    sheet.Cells[i, 4].Value = 0;
                    sheet.Cells[i, 5].Value = string.Format("{0:N2}", Usuarios.First().Meta);
                    sheet.Cells[i, 6].Value = string.Format("{0:N2}", Usuarios.First().MetaAtingimento);
                    sheet.Cells[i, 7].Value = string.Format("{0:N2}", Usuarios.Sum(x => x.Salario));
                    sheet.Cells[i, 8].Value = 0;
                    sheet.Cells[i, 9].Value = string.Format("{0:N2}", Usuarios.Sum(x => x.Comissao));
                    sheet.Cells[i, 10].Value = 0;
                    sheet.Cells[i, 11].Value = string.Format("{0:N2}", Usuarios.Sum(x => x.Total));
                    sheet.Cells[i, 12].Value = string.Format("{0:N2}", Usuarios.Sum(x => x.Teto));
                    sheet.Cells[i, 13].Value = string.Format("{0:N2}", Usuarios.Sum(x => x.Paga));
                    i++;
                }


                sheet.Cells[i, 1].Value = "Total:";
                sheet.Cells[i, 2].Value = $"{Mes}/{Ano}";
                sheet.Cells[i, 3].Value = string.Format("{0:N2}", ComercialOrtopedia.Sum(x => x.Faturado) + ComercialBuco.Sum(x => x.Faturado) + ComercialTorax.Sum(x => x.Faturado) + ComercialInterior.Sum(x => x.Faturado) + UsuariosDental.Sum(x => x.Faturado));
                sheet.Cells[i, 4].Value = string.Format("{0:N2}", UsuariosSub.FirstOrDefault()?.Faturado + UsuariosLicitacoes.Sum(x => x.Faturado));
                sheet.Cells[i, 5].Value = "";
                sheet.Cells[i, 6].Value = "";
                sheet.Cells[i, 7].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Salario) + UsuariosLicitacoes.Sum(x => x.Salario) + ComercialOrtopedia.Sum(x => x.Salario) + ComercialBuco.Sum(x => x.Salario) + ComercialTorax.Sum(x => x.Salario) + ComercialInterior.Sum(x => x.Salario) + UsuariosDental.Sum(x => x.Salario) + Usuarios.Sum(x => x.Salario));
                sheet.Cells[i, 8].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Garantia) + UsuariosLicitacoes.Sum(x => x.Garantia) + ComercialOrtopedia.Sum(x => x.Garantia) + ComercialBuco.Sum(x => x.Garantia) + ComercialTorax.Sum(x => x.Garantia) + ComercialInterior.Sum(x => x.Garantia) + UsuariosDental.Sum(x => x.Garantia));
                sheet.Cells[i, 9].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Comissao) + UsuariosLicitacoes.Sum(x => x.Comissao) + ComercialOrtopedia.Sum(x => x.Comissao) + ComercialBuco.Sum(x => x.Comissao) + ComercialTorax.Sum(x => x.Comissao) + ComercialInterior.Sum(x => x.Comissao) + UsuariosDental.Sum(x => x.Comissao) + Usuarios.Sum(x => x.Comissao));
                sheet.Cells[i, 10].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.ComissaoEquipe) + UsuariosLicitacoes.Sum(x => x.ComissaoEquipe) + ComercialOrtopedia.Sum(x => x.ComissaoEquipe) + ComercialBuco.Sum(x => x.ComissaoEquipe) + ComercialTorax.Sum(x => x.ComissaoEquipe) + ComercialInterior.Sum(x => x.ComissaoEquipe) + UsuariosDental.Sum(x => x.ComissaoEquipe) + UsuariosSub.Sum(x => x.ComissaoProduto) + UsuariosLicitacoes.Sum(x => x.ComissaoProduto) + ComercialOrtopedia.Sum(x => x.ComissaoProduto) + ComercialBuco.Sum(x => x.ComissaoProduto) + ComercialTorax.Sum(x => x.ComissaoProduto) + ComercialInterior.Sum(x => x.ComissaoProduto));
                sheet.Cells[i, 11].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Total) + UsuariosLicitacoes.Sum(x => x.Total) + ComercialOrtopedia.Sum(x => x.Total) + ComercialBuco.Sum(x => x.Total) + ComercialTorax.Sum(x => x.Total) + ComercialInterior.Sum(x => x.Total) + UsuariosDental.Sum(x => x.Total) + Usuarios.Sum(x => x.Total));
                sheet.Cells[i, 12].Value = "";
                sheet.Cells[i, 13].Value = string.Format("{0:N2}", UsuariosSub.Sum(x => x.Paga) + UsuariosLicitacoes.Sum(x => x.Paga) + ComercialOrtopedia.Sum(x => x.Paga) + ComercialBuco.Sum(x => x.Paga) + ComercialTorax.Sum(x => x.Paga) + ComercialInterior.Sum(x => x.Paga) + UsuariosDental.Sum(x => x.Paga) + Usuarios.Sum(x => x.Paga));

                var format = sheet.Cells[i, 3, i, 13];
                format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FaturamentoXFuncionamento.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoXFuncionamento Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
