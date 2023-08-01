using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models;
using SGID.Models.Account.RH;
using SGID.Models.Denuo;
using SGID.Models.Financeiro;
using SGID.Models.Inter;

namespace SGID.Pages.Account.RH
{
    [Authorize(Roles = "Admin,RH,Diretoria")]
    public class ListarTimeModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public string TextoMensagem { get; set; } = "";

        public string Mes { get; set; }

        public string MesAno { get; set; }
        public string Ano { get; set; }
        public string Empresa1 { get; set; }

        public List<Time> Users { get; set; }
        public List<TimeRH> Usuarios { get; set; } = new List<TimeRH>();

        public ListarTimeModel(ApplicationDbContext sgid,TOTVSINTERContext inter,TOTVSDENUOContext denuo)
        {
            SGID = sgid;
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
            Usuarios = new List<TimeRH>();
        }

        public void OnGet()
        {
            try
            {
                Users = SGID.Times.Where(x=>x.Status).OrderBy(x => x.Integrante).ToList();

                var date = DateTime.Now;

                MesAno = date.ToString("MM").ToUpper();
                this.Ano = date.ToString("yyyy").ToUpper();

                string data = date.ToString("yyyy/MM").Replace("/", "");
                Mes = date.ToString("MMMM").ToUpper();
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
                             join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                             where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                             && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                             && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                             && (int)(object)SD10.D1Dtdigit >= 20200701
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
                                 Linha = SA30.A3Xdescun,
                                 DOR = SB10.B1Ugrpint,
                                 SA10.A1Xgrinte
                             }
                             ).ToList();

                if (teste.Count != 0)
                {
                    teste.ForEach(x =>
                    {
                        if (!DevolucaoInter.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
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

                            DevolucaoInter.Add(new RelatorioDevolucaoFat
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
                                Login = x.A3Xlogin.Trim(),
                                Gestor = x.A3Xlogsup.Trim(),
                                Linha = x.Linha.Trim(),
                                DOR = x.DOR,
                                Codigo = x.A1Xgrinte
                            });
                        }
                    });
                }

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
                                      Vendedor = SC50.C5Nomvend,
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
                              join SA10 in ProtheusDenuo.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                              join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                              where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                              && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                              && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                              && (int)(object)SD10.D1Dtdigit >= 20200801
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
                                  Linha = SA30.A3Xdescun,
                                  SA10.A1Xgrinte
                              }
                             ).ToList();

                if (teste2.Count != 0)
                {
                    teste2.ForEach(x =>
                    {
                        if (!DevolucaoDenuo.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                        {

                            var Iguais = teste2
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

                            DevolucaoDenuo.Add(new RelatorioDevolucaoFat
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
                                Login = x.A3Xlogin.Trim(),
                                Gestor = x.A3Xlogsup.Trim(),
                                Linha = x.Linha.Trim(),
                                Codigo = x.A1Xgrinte
                            });
                        }
                    });
                }
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
                                      Vendedor = SC50.C5Nomvend,
                                      TipoCliente = SA10.A1Clinter,
                                      CodigoCliente = SA10.A1Xgrinte,
                                      Login = SA30.A3Xlogin,
                                      Gestor = SA30.A3Xlogsup
                                  }).ToList();

                #endregion

                #endregion

                var dataini = Convert.ToInt32(DataInicio);

                Users.ForEach(x =>
                {
                    var usuario = x.Integrante.ToUpper();

                    var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == x.Id).ToList();

                    var time = new TimeRH
                    {
                        User = x,
                    };



                    if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                    {
                        if(dataini < 20230801)
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
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total);
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
                            Usuarios.Add(time);
                        }
                        else
                        {
                            time.Faturado += resultadoInter.Where(x => x.Login == usuario && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                            time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);

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

                            if (usuario == "MICHEL.SAMPAIO")
                            {
                                time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total);
                                time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total);
                            }
                            else
                            {
                                time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total);
                                time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total);
                            }

                            int i = 0;
                            Produtos.ForEach(prod =>
                            {
                                if (usuario == "TIAGO.FONSECA")
                                {
                                    if (i == 0)
                                    {
                                        time.Linha = "BUCOMAXILO";
                                        time.FaturadoEquipe = 0;
                                        time.FaturadoProduto += resultadoInter.Where(x => (x.Linha == prod.Produto && x.DOR != "082") && x.Gestor != "RONAN.JOVINO" && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total) - DevolucaoInter.Where(x => (x.Linha.Trim() == prod.Produto && x.DOR != "082") && x.Gestor != "RONAN.JOVINO" && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012")) && (x.Gestor != "RONAN.JOVINO" || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012")) && (x.Gestor != "RONAN.JOVINO" || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA")).Sum(x => x.Total);
                                        i++;
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012")) && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012")) && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012")) && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012")) && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                    }
                                }
                                else if (usuario == "ARTEMIO.COSTA")
                                {
                                    time.Faturado = 0;
                                    time.FaturadoEquipe = 0;
                                    time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total);
                                    time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total);
                                }
                                else
                                {
                                    time.FaturadoEquipe = 0;
                                    time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total);
                                    time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (Convert.ToInt32(x.Data) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (Convert.ToInt32(x.Digitacao) >= 20230801 && (x.Codigo != "000011" || x.Codigo != "000012"))).Sum(x => x.Total);
                                }
                            });


                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                            time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);

                            Usuarios.Add(time);
                        }
                    }
                    else if(x.TipoFaturamento != "S")
                    {


                        time.Faturado = BaixaDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                        
                        Usuarios.Add(time);
                    }
                    else
                    {

                        time.Faturado = BaixaDenuo.Where(x=> x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                        time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                        Usuarios.Add(time);
                    }
                });
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ListarTime", user);
            }
        }

        public IActionResult OnPost(string Mes,string Ano,string Empresa)
        {
            try
            {
                Users = SGID.Times.Where(x => x.Status).OrderBy(x => x.Integrante).ToList();

                string Tempo = $"{Mes}/01/{Ano}";

                MesAno = Mes;
                this.Ano = Ano;
                Empresa1 = Empresa;

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
                             join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                             where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                             && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                             && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                             && (int)(object)SD10.D1Dtdigit >= 20200701
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
                                 Linha = SA30.A3Xdescun,
                                 DOR = SB10.B1Ugrpint,
                                 SA10.A1Xgrinte
                             }
                             ).ToList();

                if (teste.Count != 0)
                {
                    teste.ForEach(x =>
                    {
                        if (!DevolucaoInter.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
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

                            DevolucaoInter.Add(new RelatorioDevolucaoFat
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
                                Login = x.A3Xlogin.Trim(),
                                Gestor = x.A3Xlogsup.Trim(),
                                Linha = x.Linha.Trim(),
                                DOR = x.DOR,
                                Codigo = x.A1Xgrinte
                            });
                        }
                    });
                }

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
                                      Vendedor = SC50.C5Nomvend,
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
                              join SA10 in ProtheusDenuo.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                              join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                              where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                              && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                              && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                              && (int)(object)SD10.D1Dtdigit >= 20200801
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
                                  Linha = SA30.A3Xdescun,
                                  SA10.A1Xgrinte
                              }
                             ).ToList();

                if (teste2.Count != 0)
                {
                    teste2.ForEach(x =>
                    {
                        if (!DevolucaoDenuo.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                        {

                            var Iguais = teste2
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

                            DevolucaoDenuo.Add(new RelatorioDevolucaoFat
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
                                Login = x.A3Xlogin.Trim(),
                                Gestor = x.A3Xlogsup.Trim(),
                                Linha = x.Linha.Trim(),
                                Codigo = x.A1Xgrinte
                            });
                        }
                    });
                }
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
                                      Vendedor = SC50.C5Nomvend,
                                      TipoCliente = SA10.A1Clinter,
                                      CodigoCliente = SA10.A1Xgrinte,
                                      Login = SA30.A3Xlogin,
                                      Gestor = SA30.A3Xlogsup
                                  }).ToList();

                #endregion

                #endregion

                var dataini = Convert.ToInt32(DataInicio);

                Users.ForEach(x =>
                {

                    var usuario = x.Integrante.ToUpper();

                    var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == x.Id).ToList();

                    var empresa = SGID.Users.FirstOrDefault(c => c.Id == x.IdUsuario).NormalizedEmail;

                    if (Empresa == empresa.Split("@")[1]) 
                    {

                        var time = new TimeRH
                        {
                            User = x,

                        };

                        if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                        {
                            if (dataini < 20230801)
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);

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
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total);
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
                                Usuarios.Add(time);
                            }
                            else
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario  && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario  && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);

                                time.Faturado += BaixaDenuo.Where(x => x.Login == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                time.Faturado += BaixaInter.Where(x => x.Login == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);


                                time.Linha = resultadoDenuo.FirstOrDefault(x => x.Login == usuario)?.Linha;
                                if (time.Linha == null)
                                {
                                    time.Linha = resultadoInter.FirstOrDefault(x => x.Login == usuario)?.Linha;
                                }
                                else if (time.User.Integrante.ToUpper() == "EDUARDO.ARONI")
                                {
                                    time.Linha = "ORTOPEDIA";
                                }
                                else if (time.User.Integrante.ToUpper() == "TIAGO.FONSECA" || time.User.Integrante.ToUpper() == "ARTEMIO.COSTA")
                                {
                                    time.Linha = "BUCOMAXILO";
                                }

                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total);

                                    time.FaturadoEquipe += BaixaDenuo.Where(x => x.Gestor == "ANDRE.SALES" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += BaixaInter.Where(x => x.Gestor == "ANDRE.SALES" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);

                                }
                                else
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total);
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario).Sum(x => x.Total);

                                    time.FaturadoEquipe += BaixaDenuo.Where(x => x.Gestor == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += BaixaInter.Where(x => x.Gestor == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                }
                                int i = 0;
                                Produtos.ForEach(prod =>
                                {
                                    if (usuario == "TIAGO.FONSECA")
                                    {
                                        if (i == 0)
                                        {
                                            time.FaturadoEquipe = 0;
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total);

                                            time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO" || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA") && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);

                                            time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim()  && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim()  && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);

                                        time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                        time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    }
                                });
                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);
                                Usuarios.Add(time);
                            }
                        }
                        else if (x.TipoFaturamento != "S")
                        {

                            time.Faturado = BaixaDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            

                            Usuarios.Add(time);
                        }
                        else
                        {
                            time.Faturado = BaixaDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            Usuarios.Add(time);
                        }
                    }
                    else if (Empresa == null || Empresa == "")
                    {
                        var time = new TimeRH
                        {
                            User = x,

                        };

                        if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                        {
                            if (dataini < 20230801)
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);

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
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total);
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

                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                    }
                                });
                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);
                                Usuarios.Add(time);
                            }
                            else
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);

                                time.Faturado += BaixaDenuo.Where(x => x.Login == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                time.Faturado += BaixaInter.Where(x => x.Login == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);


                                time.Linha = resultadoDenuo.FirstOrDefault(x => x.Login == usuario)?.Linha;
                                if (time.Linha == null)
                                {
                                    time.Linha = resultadoInter.FirstOrDefault(x => x.Login == usuario)?.Linha;
                                }
                                else if (time.User.Integrante.ToUpper() == "EDUARDO.ARONI")
                                {
                                    time.Linha = "ORTOPEDIA";
                                }
                                else if (time.User.Integrante.ToUpper() == "TIAGO.FONSECA" || time.User.Integrante.ToUpper() == "ARTEMIO.COSTA")
                                {
                                    time.Linha = "BUCOMAXILO";
                                }

                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012") && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total);

                                    time.FaturadoEquipe += BaixaDenuo.Where(x => x.Gestor == "ANDRE.SALES" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += BaixaInter.Where(x => x.Gestor == "ANDRE.SALES" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);

                                }
                                else
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012") && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total);
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario).Sum(x => x.Total);

                                    time.FaturadoEquipe += BaixaDenuo.Where(x => x.Gestor == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += BaixaInter.Where(x => x.Gestor == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                }
                                int i = 0;
                                Produtos.ForEach(prod =>
                                {
                                    if (usuario == "TIAGO.FONSECA")
                                    {
                                        if (i == 0)
                                        {
                                            time.FaturadoEquipe = 0;
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total);

                                            time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO" || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA") && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);

                                            time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);

                                        time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                        time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    }
                                });
                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);
                                Usuarios.Add(time);
                            }
                        }
                        else if (x.TipoFaturamento != "S")
                        {
                            time.Faturado = BaixaDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);


                            Usuarios.Add(time);
                        }
                        else
                        {
                            time.Faturado = BaixaDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            Usuarios.Add(time);
                        }
                    }
                });

                return Page();
            }
            catch(Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ListarTime Post", user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(string Mes, string Ano, string Empresa)
        {
            try
            {
                Users = SGID.Times.Where(x => x.Status).OrderBy(x => x.Integrante).ToList();

                string Tempo = $"{Mes}/01/{Ano}";

                MesAno = Mes;
                this.Ano = Ano;

                var date = DateTime.Parse(Tempo);

                string data = date.ToString("yyyy/MM").Replace("/", "");
                this.Mes = date.ToString("MMMM").ToUpper();
                string DataInicio = data + "01";
                string DataFim = data + "31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109 };

                List<string> Linhas = new List<string>();
                List<RelatorioFaturamentoLinhas> LinhasValor = new List<RelatorioFaturamentoLinhas>();

                var linhasInter = ProtheusInter.Sa3010s.Where(x => x.A3Xdescun != "").Select(x => x.A3Xdescun.Trim()).Distinct().ToList();
                var linhasDenuo = ProtheusDenuo.Sa3010s.Where(x => x.A3Xdescun != "").Select(x => x.A3Xdescun.Trim()).Distinct().ToList();

                Linhas = linhasInter.Union(linhasDenuo).ToList();


                Linhas.ForEach(x =>
                {
                    LinhasValor.Add(new RelatorioFaturamentoLinhas { Nome = x });
                });


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
                             join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                             where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                             && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                             && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                             && (int)(object)SD10.D1Dtdigit >= 20200701
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
                                 Linha = SA30.A3Xdescun,
                                 DOR = SB10.B1Ugrpint,
                                 SA10.A1Xgrinte
                             }
                             ).ToList();

                if (teste.Count != 0)
                {
                    teste.ForEach(x =>
                    {
                        if (!DevolucaoInter.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
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

                            DevolucaoInter.Add(new RelatorioDevolucaoFat
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
                                Login = x.A3Xlogin.Trim(),
                                Gestor = x.A3Xlogsup.Trim(),
                                Linha = x.Linha.Trim(),
                                DOR = x.DOR,
                                Codigo = x.A1Xgrinte
                            });
                        }
                    });
                }

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
                                      Vendedor = SC50.C5Nomvend,
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
                              join SA10 in ProtheusDenuo.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                              join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                              where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                              && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                              && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                              && (int)(object)SD10.D1Dtdigit >= 20200801
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
                                  Linha = SA30.A3Xdescun,
                                  SA10.A1Xgrinte
                              }
                             ).ToList();

                if (teste2.Count != 0)
                {
                    teste2.ForEach(x =>
                    {
                        if (!DevolucaoDenuo.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                        {

                            var Iguais = teste2
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

                            DevolucaoDenuo.Add(new RelatorioDevolucaoFat
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
                                Login = x.A3Xlogin.Trim(),
                                Gestor = x.A3Xlogsup.Trim(),
                                Linha = x.Linha.Trim(),
                                Codigo = x.A1Xgrinte
                            });
                        }
                    });
                }
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
                                      Vendedor = SC50.C5Nomvend,
                                      TipoCliente = SA10.A1Clinter,
                                      CodigoCliente = SA10.A1Xgrinte,
                                      Login = SA30.A3Xlogin,
                                      Gestor = SA30.A3Xlogsup
                                  }).ToList();

                #endregion

                #endregion

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Comissoes");


                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Comissoes");

                sheet.Cells[1, 1].Value = "REPRESENTANTE";
                sheet.Cells[1, 2].Value = "FATURAMENTO MÊS";
                sheet.Cells[1, 3].Value = "META MENSAL";
                sheet.Cells[1, 4].Value = "ATINGIMENTO META";
                sheet.Cells[1, 5].Value = "COMISSÃO VENDEDOR";
                sheet.Cells[1, 6].Value = "COMISSÃO EQUIPE";
                sheet.Cells[1, 7].Value = "TOTAL BRUTO";

                var colFromHex = System.Drawing.ColorTranslator.FromHtml("#6099D5");
                var format = sheet.Cells[1, 1, 1, 7];
                format.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                format.Style.Fill.BackgroundColor.SetColor(colFromHex);
                var fontColor = System.Drawing.ColorTranslator.FromHtml("#ffffff");
                format.Style.Font.Color.SetColor(fontColor);

                int id = 2;

                var dataini = Convert.ToInt32(DataInicio);

                Users.ForEach(x =>
                {

                    var usuario = x.Integrante.ToUpper();

                    var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == x.Id).ToList();

                    var empresa = SGID.Users.FirstOrDefault(c => c.Id == x.IdUsuario).NormalizedEmail;

                    if (Empresa == empresa.Split("@")[1])
                    {

                        var time = new TimeRH
                        {
                            User = x,
                            Linha = ""
                        };

                        if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                        {
                            if (dataini < 20230801)
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);

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
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total);
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
                                Usuarios.Add(time);
                            }
                            else
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);

                                time.Faturado += BaixaDenuo.Where(x => x.Login == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                time.Faturado += BaixaInter.Where(x => x.Login == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);


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

                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total);

                                    time.FaturadoEquipe += BaixaDenuo.Where(x => x.Gestor == "ANDRE.SALES" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += BaixaInter.Where(x => x.Gestor == "ANDRE.SALES" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);

                                }
                                else
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total);
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario).Sum(x => x.Total);

                                    time.FaturadoEquipe += BaixaDenuo.Where(x => x.Gestor == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += BaixaInter.Where(x => x.Gestor == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                }
                                int i = 0;
                                Produtos.ForEach(prod =>
                                {
                                    if (usuario == "TIAGO.FONSECA")
                                    {
                                        if (i == 0)
                                        {
                                            time.FaturadoEquipe = 0;
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total);

                                            time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO" || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA") && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);

                                            time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);

                                        time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                        time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    }
                                });
                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);
                                Usuarios.Add(time);
                            }
                        }
                        else if (x.TipoFaturamento != "S")
                        {

                            time.Faturado = BaixaDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);


                            time.Linha = "LICITAÇÔES";
                            

                            Usuarios.Add(time);
                        }
                        else
                        {
                            time.Faturado = BaixaDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            time.Linha = "SUBDISTRIBUIDOR";

                            Usuarios.Add(time);
                        }

                    }
                    else if (Empresa == null || Empresa == "")
                    {
                        var time = new TimeRH
                        {
                            User = x,
                            Linha = ""
                        };

                        if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                        {
                            if (dataini < 20230801)
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);

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
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total);
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
                                Usuarios.Add(time);
                            }
                            else
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario).Sum(x => x.Total);
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario).Sum(x => x.Total);

                                time.Faturado += BaixaDenuo.Where(x => x.Login == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                time.Faturado += BaixaInter.Where(x => x.Login == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);


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

                                if (usuario == "MICHEL.SAMPAIO")
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total);

                                    time.FaturadoEquipe += BaixaDenuo.Where(x => x.Gestor == "ANDRE.SALES" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += BaixaInter.Where(x => x.Gestor == "ANDRE.SALES" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);

                                }
                                else
                                {
                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082").Sum(x => x.Total);
                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" || x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario).Sum(x => x.Total);

                                    time.FaturadoEquipe += BaixaDenuo.Where(x => x.Gestor == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += BaixaInter.Where(x => x.Gestor == usuario && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                }
                                int i = 0;
                                Produtos.ForEach(prod =>
                                {
                                    if (usuario == "TIAGO.FONSECA")
                                    {
                                        if (i == 0)
                                        {
                                            time.FaturadoEquipe = 0;
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA").Sum(x => x.Total);

                                            time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO" || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS" || x.Login == "DENIS.SOUZA") && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);

                                            time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                            time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Codigo != "000011" || x.Codigo != "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);

                                        time.FaturadoEquipe += BaixaDenuo.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                        time.FaturadoEquipe += BaixaInter.Where(x => x.Linha == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && Convert.ToInt32(x.DataPedido) >= 20230801 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012")).Sum(x => x.TotalBaixado);
                                    }
                                });
                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);
                                Usuarios.Add(time);
                            }
                        }
                        else if (x.TipoFaturamento != "S")
                        {

                            time.Faturado = BaixaDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            time.Linha = "LICITAÇÔES";

                            Usuarios.Add(time);
                        }
                        else
                        {
                            time.Faturado = BaixaDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            time.Linha = "SUBDISTRIBUIDOR";

                            Usuarios.Add(time);
                        }
                    }
                });

                LinhasValor.ForEach(x =>
                {

                    if (Usuarios.FirstOrDefault(c => c.Linha == x.Nome) != null)
                    {

                        var Faturamentos = Usuarios.Where(d => d.Linha == x.Nome).ToList();
                       
                        Faturamentos.ForEach(time =>
                        {
                            sheet.Cells[id, 1].Value = time.User.Integrante.Split("@")[0].Replace(".", " ").ToUpper();
                            if(time.User.GerenProd == "SIM" && time.User.Integrante.Split("@")[0].ToUpper() != "ARTEMIO.COSTA")
                            {
                                sheet.Cells[id, 2].Value = time.FaturadoProduto;
                            }
                            else
                            {
                                sheet.Cells[id, 2].Value = time.Faturado;
                            }

                            sheet.Cells[id, 2].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                            sheet.Cells[id, 3].Value = time.User.Meta / 12;
                            sheet.Cells[id, 3].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                            if (time.User.Meta == 0)
                            {
                                sheet.Cells[id, 4].Value = "";
                            }
                            else if (time.Faturado <= 0)
                            {
                                sheet.Cells[id, 4].Value = $" 0 %";
                            }
                            else
                            {
                                sheet.Cells[id, 4].Value = $"{string.Format("{0:0.00}", (((time.Faturado + time.FaturadoProduto) / (time.User.Meta / 12)) * 100))} %";
                            }

                            sheet.Cells[id, 5].Value = time.Comissao;
                            sheet.Cells[id, 5].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[id, 6].Value = time.ComissaoEquipe + time.ComissaoProduto;
                            sheet.Cells[id, 6].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                            sheet.Cells[id, 7].Value = "";

                            id++;
                        });

                        sheet.Cells[id, 1].Value = x.Nome;
                        sheet.Cells[id, 2].Value = Faturamentos.Sum(valor => valor.Faturado);
                        sheet.Cells[id, 2].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";



                        sheet.Cells[id, 3].Value = Faturamentos.Where(x => x.FaturadoProduto == 0).Sum(x=> x.User.Meta) / 12;
                        sheet.Cells[id, 3].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                        sheet.Cells[id, 4].Value = $"{string.Format("{0:0.00}", ((Faturamentos.Sum(valor => valor.Faturado) / (Faturamentos.Where(x=> x.FaturadoProduto == 0).Sum(x => x.User.Meta) / 12)) * 100))} %";
                        

                        sheet.Cells[id, 5].Value = Faturamentos.Sum(x => x.Comissao);
                        sheet.Cells[id, 5].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                        sheet.Cells[id, 6].Value = Faturamentos.Sum(x => x.ComissaoEquipe) + Faturamentos.Sum(x => x.ComissaoProduto);
                        sheet.Cells[id, 6].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                        sheet.Cells[id, 7].Value = "";

                        var colFromHex = System.Drawing.ColorTranslator.FromHtml("#C4E1B4");
                        
                        var format = sheet.Cells[id, 1, id, 7];
                        format.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        format.Style.Fill.BackgroundColor.SetColor(colFromHex);

                        id++;
                    }
                });

                

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
