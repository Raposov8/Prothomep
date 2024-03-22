using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Account.RH;
using SGID.Models.Denuo;
using SGID.Models.RH;

namespace SGID.Pages.Account.RH
{
    public class ListarDentalModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext Protheus { get; set; }
        public string TextoMensagem { get; set; } = "";
        public string Mes { get; set; }

        public string MesAno { get; set; }
        public string Ano { get; set; }
        public string Empresa1 { get; set; }

        public List<TimeDental> Users { get; set; }
        public List<TimeDentalRH> Usuarios { get; set; } = new List<TimeDentalRH>();

        public ListarDentalModel(ApplicationDbContext sgid, TOTVSDENUOContext denuo)
        {
            SGID = sgid;
            Protheus = denuo;
            Usuarios = new List<TimeDentalRH>();
        }
        public void OnGet()
        {
            try
            {
                Users = SGID.TimeDentals.OrderBy(x => x.Integrante).ToList();

                var date = DateTime.Now;

                MesAno = date.ToString("MM").ToUpper();
                this.Ano = date.ToString("yyyy").ToUpper();

                string data = date.ToString("yyyy/MM").Replace("/", "");
                Mes = date.ToString("MMMM").ToUpper();
                string DataInicio = data + "01";
                string DataFim = data + "31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };

                var cf = new[] { "5551", "6551", "6107", "6109" };


                #region Dental

                 #region Faturado
                    var query3 = (from SD20 in Protheus.Sd2010s
                                 join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                 join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                 join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                 join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                                 join SC50 in Protheus.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num, Cli = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                                 join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod into sr
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

                var queryDevolucao3 = (from SD10 in Protheus.Sd1010s
                                      join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Fornece = SF20.F2Cliente, Loja = SF20.F2Loja }
                                      join SD20 in Protheus.Sd2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja, Item = SD10.D1Itemori } equals new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Fornece = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Item }
                                      join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                      join SA10 in Protheus.Sa1010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                      join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
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


                #endregion

                Users.ForEach(x =>
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
                        Usuarios.Add(time);
                    }
                    else
                    {
                        time.FaturadoEquipe += Faturamento.Sum(x => x.D2_TOTAL) + Devolucao.Sum(x => x.Total);

                        time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                        Usuarios.Add(time);
                    }
                });

            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ListarDental", user);
            }
        }

        public IActionResult OnPost(string Mes, string Ano, string Empresa)
        {
            try
            {
                Users = SGID.TimeDentals.OrderBy(x => x.Integrante).ToList();

                string Tempo = $"{Mes}/01/{Ano}";

                MesAno = Mes;
                this.Ano = Ano;
                Empresa1 = Empresa;

                var date = DateTime.Parse(Tempo);

                string data = date.ToString("yyyy/MM").Replace("/", "");
                this.Mes = date.ToString("MMMM").ToUpper();
                string DataInicio = data + "01";
                string DataFim = data + "31";
                var cf = new[] { "5551", "6551", "6107", "6109" };


                #region Dental

                #region Faturado
                var query = (from SD20 in Protheus.Sd2010s
                             join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                             join SC50 in Protheus.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num, Cli = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod into sr
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

                var Faturamento = query.Select(x => new RelatorioFaturamentoDental
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
                    A3_LOGIN = x.Key.A3_LOGIN.Trim()
                }).OrderBy(x => x.A3_NOME).ToList();
                #endregion

                #region Devolucao
                var CfDevolucao = new string[] { "1202", "2202", "3202", "1553", "2553" };

                var queryDevolucao = (from SD10 in Protheus.Sd1010s
                                      join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Fornece = SF20.F2Cliente, Loja = SF20.F2Loja }
                                      join SD20 in Protheus.Sd2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja, Item = SD10.D1Itemori } equals new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Fornece = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Item }
                                      join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                      join SA10 in Protheus.Sa1010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                      join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
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

                var Devolucao = queryDevolucao.Select(x => new RelatorioDevolucaoDental
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
                    Login = x.Key.A3Xlogin.Trim()
                }).OrderBy(x => x.Vendedor).ToList();

                #endregion


                #endregion

                Users.ForEach(x =>
                {
                    var usuario = x.Integrante.ToUpper();


                    var time = new TimeDentalRH
                    {
                        User = x,
                    };

                    if (time.User.TipoComissao == "M")
                    {
                        time.Faturado += Faturamento.Where(x => x.A3_LOGIN == usuario).Sum(x => x.D2_TOTAL) + Devolucao.Where(x => x.Login == usuario).Sum(x => x.Total);

                        var valorEtapa1 = time.Meta * (time.User.AtingimentoMeta / 100);

                        if (valorEtapa1 >= time.Faturado)
                        {
                            time.Comissao = time.Faturado * (time.User.PorcentagemEtapaUm / 100);
                        }
                        else
                        {
                            time.Comissao = valorEtapa1 * (time.User.PorcentagemEtapaUm / 100);

                            time.Comissao += (time.Faturado- valorEtapa1) * (time.User.PorcentagemEtapaDois / 100);

                        }
                        
                        
                        Usuarios.Add(time);
                    }
                    else if (usuario != "MARCOS.PARRA")
                    {

                        time.Faturado += Faturamento.Where(x => x.A3_LOGIN == usuario).Sum(x => x.D2_TOTAL) + Devolucao.Where(x => x.Login == usuario).Sum(x => x.Total);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                        Usuarios.Add(time);
                    }
                    else
                    {
                        time.FaturadoEquipe += Faturamento.Sum(x => x.D2_TOTAL) + Devolucao.Sum(x => x.Total);

                        time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                        Usuarios.Add(time);
                    }
                });

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ListarDental Post", user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(string Mes, string Ano, string Empresa)
        {
            try
            {
                Users = SGID.TimeDentals.OrderBy(x => x.Integrante).ToList();

                string Tempo = $"{Mes}/01/{Ano}";

                MesAno = Mes;
                this.Ano = Ano;

                var date = DateTime.Parse(Tempo);

                string data = date.ToString("yyyy/MM").Replace("/", "");
                this.Mes = date.ToString("MMMM").ToUpper();
                string DataInicio = data + "01";
                string DataFim = data + "31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109 };


                var cf = new[] { "5551", "6551", "6107", "6109" };


                #region Dental

                #region Faturado
                var query = (from SD20 in Protheus.Sd2010s
                             join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                             join SC50 in Protheus.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num, Cli = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod into sr
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

                var Faturamento = query.Select(x => new RelatorioFaturamentoDental
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

                var queryDevolucao = (from SD10 in Protheus.Sd1010s
                                      join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Fornece = SF20.F2Cliente, Loja = SF20.F2Loja }
                                      join SD20 in Protheus.Sd2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja, Item = SD10.D1Itemori } equals new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Fornece = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Item }
                                      join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                      join SA10 in Protheus.Sa1010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                      join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
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

                var Devolucao = queryDevolucao.Select(x => new RelatorioDevolucaoDental
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


                #endregion

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Comissoes");


                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Comissoes");

                sheet.Cells[1, 1].Value = "NOME";
                sheet.Cells[1, 2].Value = "FATURAMENTO MÊS";
                sheet.Cells[1, 3].Value = "META MENSAL";
                sheet.Cells[1, 4].Value = "ATINGIMENTO META";
                sheet.Cells[1, 5].Value = "COMISSÃO";
                sheet.Cells[1, 6].Value = "TOTAL BRUTO";

                var colFromHex = System.Drawing.ColorTranslator.FromHtml("#6099D5");
                var format = sheet.Cells[1, 1, 1, 7];
                format.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                format.Style.Fill.BackgroundColor.SetColor(colFromHex);
                var fontColor = System.Drawing.ColorTranslator.FromHtml("#ffffff");
                format.Style.Font.Color.SetColor(fontColor);

                int id = 2;

                Users.ForEach(x =>
                {
                    var usuario = x.Integrante.ToUpper();


                    var time = new TimeDentalRH
                    {
                        User = x,
                    };


                    if (time.User.TipoComissao == "M")
                    {
                        time.Faturado += Faturamento.Where(x => x.A3_LOGIN == usuario).Sum(x => x.D2_TOTAL) + Devolucao.Where(x => x.Login == usuario).Sum(x => x.Total);

                        var valorEtapa1 = time.Meta * (time.User.AtingimentoMeta / 100);

                        if (valorEtapa1 >= time.Faturado)
                        {
                            time.Comissao = time.Faturado * (time.User.PorcentagemEtapaUm / 100);
                        }
                        else
                        {
                            time.Comissao = valorEtapa1 * (time.User.PorcentagemEtapaUm / 100);

                            time.Comissao += (time.Faturado - valorEtapa1) * (time.User.PorcentagemEtapaDois / 100);

                        }


                        Usuarios.Add(time);
                    }
                    else if(usuario != "MARCOS.PARRA")
                    {

                        time.Faturado += Faturamento.Where(x => x.A3_LOGIN == usuario).Sum(x => x.D2_TOTAL) + Devolucao.Where(x => x.Login == usuario).Sum(x => x.Total);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                        Usuarios.Add(time);
                    }
                    else
                    {
                        time.FaturadoEquipe += Faturamento.Sum(x => x.D2_TOTAL) + Devolucao.Sum(x => x.Total);

                        time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                        Usuarios.Add(time);
                    }
                });


                Usuarios.ForEach(x =>
                {
                    sheet.Cells[id, 1].Value = x.User.Integrante.Split("@")[0].Replace(".", " ").ToUpper();

                    if(x.User.Integrante.ToUpper() == "MARCOS.PARRA")
                    {
                        sheet.Cells[id, 2].Value = x.FaturadoEquipe;
                    }
                    else
                    {
                        sheet.Cells[id, 2].Value = x.Faturado;
                    }
                    


                    sheet.Cells[id, 2].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                    sheet.Cells[id, 3].Value = x.User.Meta / 12;
                    sheet.Cells[id, 3].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                    if (x.Faturado <= 0)
                    {
                        sheet.Cells[id, 4].Value = $" 0 %";
                    }
                    else
                    {
                        sheet.Cells[id, 4].Value = $"{string.Format("{0:0.00}", ((x.Faturado / (x.User.Meta / 12)) * 100))} %";
                    }

                    if (x.User.Integrante.ToUpper() == "MARCOS.PARRA")
                    {
                        sheet.Cells[id, 5].Value = x.ComissaoEquipe;
                    }
                    else
                    {
                        sheet.Cells[id, 5].Value = x.Comissao;
                    }

                    
                    sheet.Cells[id, 5].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                    sheet.Cells[id, 6].Value = "";

                    id++;
                });


                sheet.Cells[id, 1].Value = "DENTAL";
                sheet.Cells[id, 2].Value = Usuarios.FirstOrDefault().Faturado;
                sheet.Cells[id, 2].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";



                sheet.Cells[id, 3].Value = Usuarios.FirstOrDefault().Faturado / 12;
                sheet.Cells[id, 3].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                sheet.Cells[id, 4].Value = $"{string.Format("{0:0.00}", ((Usuarios.Sum(x=> x.Faturado) / (Usuarios.FirstOrDefault(x=> x.User.Integrante.ToUpper() == "MARCOS.PARRA").User.Meta / 12)) * 100))} %";


                sheet.Cells[id, 5].Value = Usuarios.Sum(x => x.Comissao) + Usuarios.Sum(x=> x.ComissaoEquipe);
                sheet.Cells[id, 5].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                sheet.Cells[id, 7].Value = "";

                var colFromHex2 = System.Drawing.ColorTranslator.FromHtml("#C4E1B4");

                var format2 = sheet.Cells[id, 1, id, 7];
                format2.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                format2.Style.Fill.BackgroundColor.SetColor(colFromHex2);


                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Comissoes.xlsx");
            }
            catch (SqlException e)
            {
                TextoMensagem = "Conexão GID -> Protheus cortada Tente Novamente";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ListarTime Export", user);

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ListarTime Export", user);

                return LocalRedirect("/error");
            }
        }
    }
}
