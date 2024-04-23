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
                Users = SGID.Times.Where(x => x.Status).OrderBy(x => x.Integrante).ToList();

                var date = DateTime.Now;

                MesAno = date.ToString("MM").ToUpper();
                this.Ano = date.ToString("yyyy").ToUpper();

                string data = date.ToString("yyyy/MM").Replace("/", "");
                Mes = date.ToString("MMMM").ToUpper();

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
                }).OrderBy(x=> x.Login).ToList();
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
                                     join SC50 in ProtheusInter.Sc5010s on new { Filial = SE10.E1Filial, Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num } into sr
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
                                            join SC50 in ProtheusInter.Sc5010s on new { Filial = SE10.E1Filial, Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num }
                                            join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                            join SD20 in ProtheusInter.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                            where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R" && SC50.DELET != "*" && SD20.DELET != "*"
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
                                                DataPedido = SD20.D2Emissao,
                                                Empresa = "INTERMEDIC"
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
                                                x.DataPedido,
                                                x.Empresa
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
                                                DataPedido = x.Key.DataPedido,
                                                Empresa = x.Key.Empresa
                                            }).ToList();

                #endregion

                #region BaixaKamikaze

                var RelatorioKamikazeInter = (from SE50 in ProtheusInter.Se5010s
                             join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                             equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                             join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                             join SC50 in ProtheusInter.Sc5010s on new { Filial = SE10.E1Filial, Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num }
                             join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             join SD20 in ProtheusInter.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                             where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R" && SA10.DELET != "*" && SA10.A1Msblql != "1"
                             && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                             && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                             && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                             && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                             && (int)(object)SE50.E5Data <= (int)(object)DataFim
                             && SC50.C5Utpoper == "K"
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
                                 Login = x.Key.Login.Trim(),
                                 Gestor = x.Key.Gestor,
                                 DataPedido = x.Key.DataPedido
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
                              where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
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
                                     join SE10 in ProtheusDenuo.Se1010s on new { Filial=SE50.E5Filial , PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                     equals new { Filial = SE10.E1Filial, PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                     join SA10 in ProtheusDenuo.Sa1010s on  SE50.E5Cliente equals SA10.A1Cod
                                     join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SE10.E1Filial, Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num } into sr
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
                                              join SE10 in ProtheusDenuo.Se1010s on new { Filial=SE50.E5Filial, PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                              equals new { Filial = SE10.E1Filial, PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                              join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                              join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SE10.E1Filial,Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num }
                                              join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                              join SD20 in ProtheusDenuo.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                              where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R" && SC50.DELET != "*" && SD20.DELET != "*"
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
                                                  DataPedido = SD20.D2Emissao,
                                                  Empresa = "DENUO"
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
                                                  x.DataPedido,
                                                  x.Empresa
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
                                                  DataPedido = x.Key.DataPedido,
                                                  Empresa = x.Key.Empresa
                                              }).ToList();

                #endregion

                #region BaixaKamikaze

                var RelatorioKamikazeDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                              join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                              equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                              join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                              join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SE10.E1Filial, Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num }
                                              join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                              join SD20 in ProtheusDenuo.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                              where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R" && SA10.DELET != "*" && SA10.A1Msblql != "1"
                                              && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                              && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                              && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                              && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                              && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                              && SC50.C5Utpoper == "K"
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
                                                  Login = x.Key.Login.Trim(),
                                                  Gestor = x.Key.Gestor,
                                                  DataPedido = x.Key.DataPedido
                                              }).ToList();

                #endregion

                #endregion

                #endregion


                var dataini = Convert.ToInt32(DataInicio);

                Users = SGID.Times.Where(x => x.Status).ToList();
                Users.AddRange(SGID.Times.Where(x=> x.Desativar > date).ToList());
                Users.ForEach(x =>
                {

                    if (Empresa != null && Empresa != "")
                    {
                        if (Empresa == "INTERMEDIC")
                        {
                            var usuario = x.Integrante.ToUpper();

                            var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == x.Id).ToList();

                            var time = new TimeRH
                            {
                                User = x,
                                Linha = ""
                            };

                            if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                            {
                               
                                time.Linha = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Xlogin == usuario)?.A3Xdescun;
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
                                    time.Linha = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Xlogin == usuario)?.A3Xdescun;
                                }

                                if (dataini <=  20240231)
                                {
                                    #region RegrasAntigas
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
                                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082").Sum(x => x.Total);
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082").Sum(x => x.Total);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        }
                                    });
                                    #endregion
                                }
                                else
                                {
                                    #region RegrasNovas
                                    time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                    time.Faturado += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                    time.Faturado += RelatorioKamikazeInter.Where(x=> x.Login == usuario).Sum(x => x.TotalBaixado);

                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += RelatorioKamikazeInter.Where(x => x.Gestor == usuario).Sum(x => x.TotalBaixado);

                                    int i = 0;
                                    Produtos.ForEach(prod =>
                                    {
                                        if (usuario == "TIAGO.FONSECA")
                                        {
                                            if (i == 0)
                                            {
                                                time.FaturadoEquipe = 0;
                                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        }
                                    });
                                    #endregion
                                }

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);
                                if (time.Faturado > 0 || time.FaturadoEquipe > 0 || time.FaturadoProduto > 0)
                                {
                                    Usuarios.Add(time);
                                }
                                
                            }
                            else if (x.TipoFaturamento != "S")
                            {

                                time.Faturado = BaixaLicitacoesInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);


                                time.Linha = "LICITAÇÔES";


                                if (time.Faturado > 0)
                                {
                                    Usuarios.Add(time);
                                }
                            }
                            else
                            {
                                time.Faturado = BaixaSubInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.Linha = "SUBDISTRIBUIDOR";

                                if(time.Faturado > 0)
                                {
                                    Usuarios.Add(time);
                                }

                            }
                        }
                        else
                        {
                            var usuario = x.Integrante.ToUpper();

                            var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == x.Id).ToList();

                            var time = new TimeRH
                            {
                                User = x,
                                Linha = ""
                            };

                            if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                            {
                                if (dataini <= 20240231)
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
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total);
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        }
                                    });
                                }
                                else
                                {
                                    time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                    time.Faturado += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                    time.Faturado += RelatorioKamikazeDenuo.Where(x => x.Login == usuario).Sum(x => x.TotalBaixado);

                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += RelatorioKamikazeDenuo.Where(x => x.Gestor == usuario).Sum(x => x.TotalBaixado);

                                    int i = 0;
                                    Produtos.ForEach(prod =>
                                    {
                                        if (usuario == "TIAGO.FONSECA")
                                        {
                                            if (i == 0)
                                            {
                                                time.FaturadoEquipe = 0;
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto  && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        }
                                    });
                                }

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);
                            }
                            else if (x.TipoFaturamento != "S")
                            {

                                time.Faturado = BaixaLicitacoesDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                                time.Linha = "LICITAÇÔES";

                                if (time.Faturado > 0)
                                {
                                    Usuarios.Add(time);
                                }
                            }
                            else
                            {
                                time.Faturado = BaixaSubDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.Linha = "SUBDISTRIBUIDOR";

                                if (time.Faturado > 0)
                                {
                                    Usuarios.Add(time);
                                }
                            }
                        }

                    }
                    else
                    {
                        var usuario = x.Integrante.ToUpper();

                        var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == x.Id).ToList();

                        var time = new TimeRH
                        {
                            User = x,
                        };

                        if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                        {
                            time.Linha = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Xlogin == usuario)?.A3Xdescun;
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
                                time.Linha = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Xlogin == usuario)?.A3Xdescun;
                            }

                            if (dataini <= 20240231)
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
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                    }
                                });
                            }
                            else
                            {
                                if(usuario == "FABIANA.MACCHIA")
                                {
                                    var teste = "";
                                }

                                time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                time.Faturado += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                time.Faturado += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                time.Faturado += RelatorioKamikazeDenuo.Where(x => x.Login == usuario).Sum(x => x.TotalBaixado);
                                time.Faturado += RelatorioKamikazeInter.Where(x => x.Login == usuario).Sum(x => x.TotalBaixado);

                                time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                time.FaturadoEquipe += RelatorioKamikazeDenuo.Where(x => x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                time.FaturadoEquipe += RelatorioKamikazeInter.Where(x => x.Gestor == usuario).Sum(x => x.TotalBaixado);

                                int i = 0;
                                Produtos.ForEach(prod =>
                                {
                                    if (usuario == "TIAGO.FONSECA")
                                    {
                                        if (i == 0)
                                        {
                                            time.FaturadoEquipe = 0;
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                    }
                                });
                                
                            }

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                            time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                            time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);
                            Usuarios.Add(time);

                        }
                        else if (x.TipoFaturamento != "S")
                        {


                            time.Faturado = BaixaLicitacoesDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaLicitacoesInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            Usuarios.Add(time);
                        }
                        else
                        {

                            time.Faturado = BaixaSubDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaSubInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            Usuarios.Add(time);
                        }
                    }
                    
                });

                Users = Users.OrderBy(x => x.Integrante).ToList();
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

                string Tempo = $"{Mes}/01/{Ano}";

                MesAno = Mes;
                this.Ano = Ano;

                var date = DateTime.Parse(Tempo);

                string data = date.ToString("yyyy/MM").Replace("/", "");
                this.Mes = date.ToString("MMMM").ToUpper();
                string DataInicio = data + "01";
                string DataFim = data + "31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };

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
                             where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SC60.DELET != "*"
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
                }).OrderBy(x => x.Login).ToList();
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
                                     join SC50 in ProtheusInter.Sc5010s on new { Filial = SE10.E1Filial, Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num } into sr
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
                                            join SC50 in ProtheusInter.Sc5010s on new { Filial = SE10.E1Filial, Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num }
                                            join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                            join SD20 in ProtheusInter.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                            where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R" && SC50.DELET != "*" && SD20.DELET != "*"
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
                                                DataPedido = SD20.D2Emissao,
                                                Empresa = "INTERMEDIC"
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
                                                x.DataPedido,
                                                x.Empresa
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
                                                DataPedido = x.Key.DataPedido,
                                                Empresa = x.Key.Empresa
                                            }).ToList();

                #endregion

                #region BaixaKamikaze

                var RelatorioKamikazeInter = (from SE50 in ProtheusInter.Se5010s
                                              join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                              equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                              join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                              join SC50 in ProtheusInter.Sc5010s on new { Filial = SE10.E1Filial, Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num }
                                              join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                              join SD20 in ProtheusInter.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                              where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R" && SA10.DELET != "*" && SA10.A1Msblql != "1"
                                              && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                              && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                              && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                              && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                              && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                              && SC50.C5Utpoper == "K"
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
                                     join SE10 in ProtheusDenuo.Se1010s on new { Filial = SE50.E5Filial, PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                     equals new { Filial = SE10.E1Filial, PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                     join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                     join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SE10.E1Filial, Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num } into sr
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
                                            join SE10 in ProtheusDenuo.Se1010s on new { Filial = SE50.E5Filial, PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                            equals new { Filial = SE10.E1Filial, PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                            join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                            join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SE10.E1Filial, Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num }
                                            join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                            join SD20 in ProtheusDenuo.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                            where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R" && SC50.DELET != "*" && SD20.DELET != "*"
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
                                                DataPedido = SD20.D2Emissao,
                                                Empresa = "DENUO"
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
                                                x.DataPedido,
                                                x.Empresa
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
                                                DataPedido = x.Key.DataPedido,
                                                Empresa = x.Key.Empresa
                                            }).ToList();

                #endregion

                #region BaixaKamikaze

                var RelatorioKamikazeDenuo = (from SE50 in ProtheusDenuo.Se5010s
                                              join SE10 in ProtheusDenuo.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                                              equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                                              join SA10 in ProtheusDenuo.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                                              join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SE10.E1Filial, Pedido = SE10.E1Pedido } equals new { Filial = SC50.C5Filial, Pedido = SC50.C5Num }
                                              join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                              join SD20 in ProtheusDenuo.Sd2010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SD20.D2Filial, Num = SD20.D2Pedido }
                                              where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R" && SA10.DELET != "*" && SA10.A1Msblql != "1"
                                              && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                                              && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                                              && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                                              && (int)(object)SE50.E5Data >= (int)(object)DataInicio
                                              && (int)(object)SE50.E5Data <= (int)(object)DataFim
                                              && SC50.C5Utpoper == "K"
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

                Users = SGID.Times.Where(x => x.Status).ToList();
                Users.AddRange(SGID.Times.Where(x => x.Desativar > date).ToList());
                Users.ForEach(x =>
                {

                    if (Empresa != null && Empresa != "")
                    {
                        if (Empresa == "INTERMEDIC")
                        {
                            var usuario = x.Integrante.ToUpper();

                            var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == x.Id).ToList();

                            var time = new TimeRH
                            {
                                User = x,
                                Linha = ""
                            };

                            if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                            {

                                time.Linha = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Xlogin == usuario)?.A3Xdescun;
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
                                    time.Linha = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Xlogin == usuario)?.A3Xdescun;
                                }

                                if (dataini <=  20240231)
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
                                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082").Sum(x => x.Total);
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082").Sum(x => x.Total);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        }
                                    });
                                }
                                else
                                {
                                    time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                    time.Faturado += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                    time.Faturado += RelatorioKamikazeInter.Where(x => x.Login == usuario).Sum(x => x.TotalBaixado);

                                    time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += RelatorioKamikazeInter.Where(x => x.Gestor == usuario).Sum(x => x.TotalBaixado);

                                    int i = 0;
                                    Produtos.ForEach(prod =>
                                    {
                                        if (usuario == "TIAGO.FONSECA")
                                        {
                                            if (i == 0)
                                            {
                                                time.FaturadoEquipe = 0;
                                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        }
                                    });
                                }

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);
                                if (time.Faturado > 0 || time.FaturadoEquipe > 0 || time.FaturadoProduto > 0)
                                {
                                    Usuarios.Add(time);
                                }

                            }
                            else if (x.TipoFaturamento != "S")
                            {

                                time.Faturado = BaixaLicitacoesInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);


                                time.Linha = "LICITAÇÔES";


                                if (time.Faturado > 0)
                                {
                                    Usuarios.Add(time);
                                }
                            }
                            else
                            {
                                time.Faturado = BaixaSubInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.Linha = "SUBDISTRIBUIDOR";

                                if (time.Faturado > 0)
                                {
                                    Usuarios.Add(time);
                                }

                            }
                        }
                        else
                        {
                            var usuario = x.Integrante.ToUpper();

                            var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == x.Id).ToList();

                            var time = new TimeRH
                            {
                                User = x,
                                Linha = ""
                            };

                            if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                            {
                                if (dataini <=  20240231)
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
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total);
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        }
                                    });
                                }
                                else
                                {
                                    time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                    time.Faturado += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20231231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                    time.Faturado += RelatorioKamikazeDenuo.Where(x => x.Login == usuario).Sum(x => x.TotalBaixado);

                                    time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario  && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                    time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20231231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                    time.FaturadoEquipe += RelatorioKamikazeDenuo.Where(x => x.Gestor == usuario).Sum(x => x.TotalBaixado);

                                    int i = 0;
                                    Produtos.ForEach(prod =>
                                    {
                                        if (usuario == "TIAGO.FONSECA")
                                        {
                                            if (i == 0)
                                            {
                                                time.FaturadoEquipe = 0;
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                i++;
                                            }
                                            else
                                            {
                                                time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                                time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                                time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            }
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20231231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        }
                                    });

                                    

                                    time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                    time.ComissaoEquipe = time.FaturadoEquipe * (time.User.PorcentagemSeg / 100);
                                    time.ComissaoProduto = time.FaturadoProduto * (time.User.PorcentagemGenProd / 100);
                                    if (time.Faturado > 0 || time.FaturadoEquipe > 0 || time.FaturadoProduto > 0)
                                    {
                                        Usuarios.Add(time);
                                    }
                                }
                            }
                            else if (x.TipoFaturamento != "S")
                            {

                                time.Faturado = BaixaLicitacoesDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                                time.Linha = "LICITAÇÔES";

                                if (time.Faturado > 0)
                                {
                                    Usuarios.Add(time);
                                }
                            }
                            else
                            {
                                time.Faturado = BaixaSubDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                                time.Comissao = time.Faturado * (time.User.Porcentagem / 100);
                                time.Linha = "SUBDISTRIBUIDOR";

                                if (time.Faturado > 0)
                                {
                                    Usuarios.Add(time);
                                }
                            }
                        }

                    }
                    else
                    {
                        var usuario = x.Integrante.ToUpper();

                        var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == x.Id).ToList();

                        var time = new TimeRH
                        {
                            User = x,
                        };



                        if (x.TipoFaturamento != "S" && x.TipoFaturamento != "L")
                        {
                            time.Linha = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Xlogin == usuario)?.A3Xdescun;
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
                                time.Linha = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Xlogin == usuario)?.A3Xdescun;
                            }

                            if (dataini <= 20240231)
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
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.Total);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                                    }
                                });
                            }
                            else
                            {
                                time.Faturado += resultadoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                time.Faturado += resultadoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                time.Faturado += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                time.Faturado += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" || x.CodigoCliente == "000012") && x.Login == usuario).Sum(x => x.TotalBaixado);
                                time.Faturado += RelatorioKamikazeDenuo.Where(x => x.Login == usuario).Sum(x => x.TotalBaixado);
                                time.Faturado += RelatorioKamikazeInter.Where(x => x.Login == usuario).Sum(x => x.TotalBaixado);


                                time.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                time.FaturadoEquipe += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                time.FaturadoEquipe += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                time.FaturadoEquipe += RelatorioKamikazeDenuo.Where(x => x.Gestor == usuario).Sum(x => x.TotalBaixado);
                                time.FaturadoEquipe += RelatorioKamikazeInter.Where(x => x.Gestor == usuario).Sum(x => x.TotalBaixado);

                                int i = 0;
                                Produtos.ForEach(prod =>
                                {
                                    if (usuario == "TIAGO.FONSECA")
                                    {
                                        if (i == 0)
                                        {
                                            time.FaturadoEquipe = 0;
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            i++;
                                        }
                                        else
                                        {
                                            time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                            time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                            time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        }
                                    }
                                    else
                                    {
                                        time.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha == prod.Produto && x.DOR != "082" && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += resultadoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && (x.Codigo != "000011" && x.Codigo != "000012")).Sum(x => x.Total);
                                        time.FaturadoProduto += BaixaLicitacoesDenuo.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        time.FaturadoProduto += BaixaLicitacoesInter.Where(x => Convert.ToInt32(x.DataPedido) > 20240231 && (x.CodigoCliente == "000011" && x.CodigoCliente == "000012") && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA" && x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        time.FaturadoProduto += RelatorioKamikazeDenuo.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                        time.FaturadoProduto += RelatorioKamikazeInter.Where(x => x.Linha == prod.Produto).Sum(x => x.TotalBaixado);
                                    }
                                });

                            }

                            Usuarios.Add(time);
                        }
                        else if (x.TipoFaturamento != "S")
                        {


                            time.Faturado = BaixaLicitacoesDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaLicitacoesInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            Usuarios.Add(time);
                        }
                        else
                        {

                            time.Faturado = BaixaSubDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaSubInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                            time.Comissao = time.Faturado * (time.User.Porcentagem / 100);

                            Usuarios.Add(time);
                        }
                    }

                });

                Usuarios = Usuarios.OrderBy(x => x.User.Integrante).ToList();

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
