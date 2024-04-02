using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models;
using SGID.Models.Account.RH;
using SGID.Models.Denuo;
using SGID.Models.Financeiro;
using SGID.Models.Controladoria.FaturamentoNF;

namespace SGID.Pages.Account.RH
{
    [Authorize(Roles = "Admin,RH,Diretoria")]
    public class ListarADMModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public string TextoMensagem { get; set; } = "";

        public string Mes { get; set; }

        public string MesAno { get; set; }
        public string Ano { get; set; }
        public string Empresa1 { get; set; }

        public List<TimeADM> Users { get; set; }
        public List<TimeADMRH> Usuarios { get; set; } = new List<TimeADMRH>();

        public ListarADMModel(ApplicationDbContext sgid, TOTVSINTERContext inter, TOTVSDENUOContext denuo)
        {
            SGID = sgid;
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
            Usuarios = new List<TimeADMRH>();
        }
        public void OnGet()
        {
            try
            {
                var date = DateTime.Now;

                MesAno = date.ToString("MM").ToUpper();
                this.Ano = date.ToString("yyyy").ToUpper();

                string data = date.ToString("yyyy/MM").Replace("/", "");
                Mes = date.ToString("MMMM").ToUpper();
                string DataInicio = data + "01";
                string DataFim = data + "31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };

                #region INTERMEDIC

                #region Faturado
                var query = (from SD20 in ProtheusInter.Sd2010s
                             join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             join SC60 in ProtheusInter.Sc6010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Pedido = SC60.C6Num, Cliente = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                             where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SC60.DELET != "*"
                             && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                             ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                             //&& SD20.D2Quant != 0 
                             && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
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
                             where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
                             && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                             && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                             && (int)(object)SD10.D1Dtdigit >= 20200701 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K")
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
                                 SA10.A1Xgrinte
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
                                 x.DOR,
                                 x.A1Xgrinte
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
                    Linha = x.Key.A3Xdescun,
                    DOR = x.Key.DOR,
                    Codigo = x.Key.A1Xgrinte
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
                              where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                 && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                 (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                 //&& SD20.D2Quant != 0 
                                 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SC50.C5Xtipopv != "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
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
                                  SA30.A3Xlogin,
                                  SA30.A3Xlogsup,
                                  SD10.D1Nfori,
                                  SD10.D1Seriori,
                                  SD10.D1Datori,
                                  SD10.D1Emissao,
                                  SA30.A3Xdescun,
                                  SA10.A1Xgrinte
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
                                 x.A1Xgrinte
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
                    Linha = x.Key.A3Xdescun,
                    Codigo = x.Key.A1Xgrinte
                }).ToList();
                #endregion

                #endregion

                Users = SGID.TimeADMs.Where(x=> x.Status).ToList();
                Users.AddRange(SGID.TimeADMs.Where(x=> x.Desativar > date).ToList());

                Users.ForEach(x =>
                {
                    var usuario = x.Integrante.ToUpper();


                    var time = new TimeADMRH
                    {
                        User = x,
                    };

                    time.Faturado += resultadoInter.Sum(x => x.Total) - DevolucaoInter.Sum(x => x.Total);
                    time.Faturado += resultadoDenuo.Sum(x => x.Total) - DevolucaoDenuo.Sum(x => x.Total);

                    time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                    Usuarios.Add(time);
                });

                Users = Users.OrderBy(x => x.Integrante).ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ListarADM", user);
            }
        }

        public IActionResult OnPost(string Mes, string Ano, string Empresa)
        {
            try
            {
                string Tempo = $"{Mes}/01/{Ano}";

                MesAno = Mes;
                this.Ano = Ano;
                Empresa1 = Empresa;

                var date = DateTime.Parse(Tempo);

                string data = date.ToString("yyyy/MM").Replace("/", "");
                this.Mes = date.ToString("MMMM").ToUpper();
                string DataInicio = data + "01";
                string DataFim = data + "31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };

                #region INTERMEDIC

                #region Faturado
                var query = (from SD20 in ProtheusInter.Sd2010s
                             join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             join SC60 in ProtheusInter.Sc6010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Pedido = SC60.C6Num, Cliente = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                             where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SC60.DELET != "*"
                             && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                             ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                             //&& SD20.D2Quant != 0 
                             && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
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
                             where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
                             && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                             && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                             && (int)(object)SD10.D1Dtdigit >= 20200701 && SC50.C5Utpoper == "F"
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
                                 SA10.A1Xgrinte
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
                                 x.DOR,
                                 x.A1Xgrinte
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
                    Linha = x.Key.A3Xdescun,
                    DOR = x.Key.DOR,
                    Codigo = x.Key.A1Xgrinte
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
                              where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                 && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                 (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                 //&& SD20.D2Quant != 0 
                                 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SC50.C5Xtipopv != "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
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
                              where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
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
                                  SA10.A1Xgrinte
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
                                 x.A1Xgrinte
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
                    Linha = x.Key.A3Xdescun,
                    Codigo = x.Key.A1Xgrinte
                }).ToList();
                #endregion

                #endregion

                Users = SGID.TimeADMs.Where(x => x.Status).ToList();
                Users.AddRange(SGID.TimeADMs.Where(x => x.Desativar > date).ToList());

                Users.ForEach(x =>
                {
                    var usuario = x.Integrante.ToUpper();


                    var time = new TimeADMRH
                    {
                        User = x,
                    };

                    time.Faturado += resultadoInter.Sum(x => x.Total) - DevolucaoInter.Sum(x => x.Total);
                    time.Faturado += resultadoDenuo.Sum(x => x.Total) - DevolucaoDenuo.Sum(x => x.Total);

                    time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                    Usuarios.Add(time);
                });

                Users = Users.OrderBy(x => x.Integrante).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ListarADM Post", user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(string Mes, string Ano, string Empresa)
        {
            try
            {
                string Tempo = $"{Mes}/01/{Ano}";

                MesAno = Mes;
                this.Ano = Ano;

                var date = DateTime.Parse(Tempo);

                string data = date.ToString("yyyy/MM").Replace("/", "");
                this.Mes = date.ToString("MMMM").ToUpper();
                string DataInicio = data + "01";
                string DataFim = data + "31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109 };


                #region INTERMEDIC

                #region Faturado
                var query = (from SD20 in ProtheusInter.Sd2010s
                             join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             join SC60 in ProtheusInter.Sc6010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Pedido = SC60.C6Num, Cliente = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                             where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SC60.DELET != "*"
                             && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                             ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                             //&& SD20.D2Quant != 0 
                             && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
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
                             where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
                             && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                             && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                             && (int)(object)SD10.D1Dtdigit >= 20200701 && SC50.C5Utpoper == "F"
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
                                 SA10.A1Xgrinte
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
                                 x.DOR,
                                 x.A1Xgrinte
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
                    Linha = x.Key.A3Xdescun,
                    DOR = x.Key.DOR,
                    Codigo = x.Key.A1Xgrinte
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
                              where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                                 && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                                 (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                                 //&& SD20.D2Quant != 0 
                                 && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SC50.C5Xtipopv != "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
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
                              where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
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
                                  SA10.A1Xgrinte
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
                                 x.A1Xgrinte
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
                    Linha = x.Key.A3Xdescun,
                    Codigo = x.Key.A1Xgrinte
                }).ToList();
                #endregion

                #endregion



                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Comissoes");


                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Comissoes");

                sheet.Cells[1, 1].Value = "ADM";
                sheet.Cells[1, 2].Value = "FATURAMENTO MÊS";
                sheet.Cells[1, 3].Value = "META MENSAL";
                sheet.Cells[1, 4].Value = "ATINGIMENTO META";
                sheet.Cells[1, 5].Value = "COMISSÃO ADM";
                sheet.Cells[1, 6].Value = "TOTAL BRUTO";

                var colFromHex = System.Drawing.ColorTranslator.FromHtml("#6099D5");
                var format = sheet.Cells[1, 1, 1, 7];
                format.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                format.Style.Fill.BackgroundColor.SetColor(colFromHex);
                var fontColor = System.Drawing.ColorTranslator.FromHtml("#ffffff");
                format.Style.Font.Color.SetColor(fontColor);

                int id = 2;

                Users = SGID.TimeADMs.Where(x => x.Status).ToList();
                Users.AddRange(SGID.TimeADMs.Where(x => x.Desativar > date).ToList());

                Users.ForEach(x =>
                {
                    var usuario = x.Integrante.ToUpper();

                    var time = new TimeADMRH
                    {
                        User = x,
                    };

                    time.Faturado += resultadoInter.Sum(x => x.Total) - DevolucaoInter.Sum(x => x.Total);
                    time.Faturado += resultadoDenuo.Sum(x => x.Total) - DevolucaoDenuo.Sum(x => x.Total);

                    time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                    Usuarios.Add(time);
                });

                Usuarios = Usuarios.OrderBy(x => x.User.Integrante).ToList();

                Usuarios.ForEach(x =>
                {
                    sheet.Cells[id, 1].Value = x.User.Integrante.Split("@")[0].Replace(".", " ").ToUpper();

                    sheet.Cells[id, 2].Value = x.Faturado;


                    sheet.Cells[id, 2].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                    sheet.Cells[id, 3].Value = 5500000;
                    sheet.Cells[id, 3].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                    if (x.Faturado <= 0)
                    {
                        sheet.Cells[id, 4].Value = $" 0 %";
                    }
                    else
                    {
                        sheet.Cells[id, 4].Value = $"{string.Format("{0:0.00}", ((x.Faturado / 5500000) * 100))} %";
                    }

                    sheet.Cells[id, 5].Value = x.Comissao;
                    sheet.Cells[id, 5].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                    sheet.Cells[id, 6].Value = "";

                    id++;
                });
                
                    
                sheet.Cells[id, 1].Value = "ADM";
                sheet.Cells[id, 2].Value = Usuarios.FirstOrDefault().Faturado;
                sheet.Cells[id, 2].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";



                sheet.Cells[id, 3].Value = Usuarios.FirstOrDefault().Faturado / 12;
                sheet.Cells[id, 3].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                sheet.Cells[id, 4].Value = $"{string.Format("{0:0.00}", ((Usuarios.FirstOrDefault().Faturado / 5500000 ) * 100))} %";


                sheet.Cells[id, 5].Value = Usuarios.Sum(x => x.Comissao);
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
