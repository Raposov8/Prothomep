using DocumentFormat.OpenXml.Vml.Office;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Models;
using SGID.Models.Account.RH;
using SGID.Models.Controladoria.FaturamentoNF;
using SGID.Models.Denuo;
using SGID.Models.Estoque.RelatorioFaturamentoNFFab;
using System.Linq;
using SGID.Models.Financeiro;

namespace SGID.Pages.Account.RH
{
    public class RelatorioComissaoModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public string MesAno { get; set; }
        public string Ano { get; set; }
        public int IdTime { get; set; }

        public string Mes { get; set; }

        public TimeRH Representante { get; set; }

        public List<RelatorioFaturamentoTime> Faturamento { get; set; } = new List<RelatorioFaturamentoTime>();
        public List<RelatorioFaturamentoTime> GestorFaturamento { get; set; } = new List<RelatorioFaturamentoTime>();
        public List<RelatorioFaturamentoTime> LinhaFaturamento { get; set; } = new List<RelatorioFaturamentoTime>();
        public List<RelatorioAreceberBaixa> BaixaFaturamento { get; set; } = new List<RelatorioAreceberBaixa>();

        public List<RelatorioDevolucaoFat> Devolucao { get; set; } = new List<RelatorioDevolucaoFat>();
        public List<RelatorioDevolucaoFat> GestorDevolucao { get; set; } = new List<RelatorioDevolucaoFat>();
        public List<RelatorioDevolucaoFat> LinhaDevolucao { get; set; } = new List<RelatorioDevolucaoFat>();

        public RelatorioComissaoModel(ApplicationDbContext sgid,TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            SGID = sgid;
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
        }

        public void OnGet(string Mes,string Ano,int IdTime)
        {
            this.Ano = Ano;
            MesAno = Mes;
            this.IdTime = IdTime;

            string Tempo = $"{Mes}/01/{Ano}";

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
            }).Select(x => new RelatorioFaturamentoTime
            {
                NF = x.Key.NF,
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
                              join SA30 in ProtheusInter.Sa3010s on SA10.A1Vend equals SA30.A3Cod
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
            }).Select(x => new RelatorioFaturamentoTime
            {
                NF = x.Key.NF,
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
                              join SA30 in ProtheusDenuo.Sa3010s on SA10.A1Vend equals SA30.A3Cod
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

            var usuario = SGID.Times.FirstOrDefault(x=> x.Id == IdTime);

            var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == usuario.Id).ToList();

            Representante = new TimeRH
            {
                User = usuario
            };

            if (Representante.User.TipoFaturamento != "S" && Representante.User.TipoFaturamento != "L")
            {
                Representante.Faturado += resultadoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total);

                Representante.Faturado += resultadoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total);

                Faturamento.AddRange(resultadoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
                Faturamento.AddRange(resultadoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
                Devolucao.AddRange(DevolucaoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
                Devolucao.AddRange(DevolucaoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());

                if (usuario.Integrante.ToUpper() == "MICHEL.SAMPAIO")
                {
                    Representante.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                    Representante.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                    GestorFaturamento.AddRange(resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").ToList());
                    GestorFaturamento.AddRange(resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").ToList());
                    GestorDevolucao.AddRange(DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").ToList());
                    GestorDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").ToList());
                }
                else
                {
                    Representante.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper() && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper() && x.DOR != "082").Sum(x => x.Total);
                    Representante.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper().ToUpper()).Sum(x => x.Total);
                    GestorFaturamento.AddRange(resultadoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                    GestorFaturamento.AddRange(resultadoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                    GestorDevolucao.AddRange(DevolucaoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                    GestorDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                }


                Produtos.ForEach(prod =>
                {
                    int i = 0;
                    if (usuario.Integrante.ToUpper() == "TIAGO.FONSECA")
                    {
                        if (i == 0)
                        {
                            Representante.FaturadoProduto += resultadoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total) - DevolucaoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total);
                            Representante.FaturadoProduto += resultadoInter.Where(x => (x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total) - DevolucaoInter.Where(x => (x.Linha.Trim() == prod.Produto && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total);
                            LinhaFaturamento.AddRange(resultadoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").ToList());
                            LinhaDevolucao.AddRange(DevolucaoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").ToList());
                            i++;
                        }
                        else
                        {
                            Representante.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                            Representante.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                            LinhaFaturamento.AddRange(resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").ToList());
                            LinhaDevolucao.AddRange(DevolucaoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").ToList());
                            LinhaFaturamento.AddRange(resultadoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO")).ToList());
                            LinhaDevolucao.AddRange(DevolucaoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO")).ToList());
                        }
                    }
                    else
                    {
                        Representante.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                        Representante.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                        LinhaFaturamento.AddRange(resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                        LinhaDevolucao.AddRange(DevolucaoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                        LinhaFaturamento.AddRange(resultadoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                        LinhaDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                    }
                });


                Representante.Comissao = Representante.Faturado * (Representante.User.Porcentagem / 100);
                Representante.ComissaoEquipe = Representante.FaturadoEquipe * (Representante.User.PorcentagemSeg / 100);
                Representante.ComissaoProduto = Representante.FaturadoProduto * (Representante.User.PorcentagemGenProd / 100);
            }
            else if (Representante.User.TipoFaturamento != "S")
            {


                Representante.Faturado = BaixaDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                Representante.Comissao = Representante.Faturado * (Representante.User.Porcentagem / 100);

                BaixaFaturamento.AddRange(BaixaDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").ToList());
            }
            else
            {

                Representante.Faturado = BaixaDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                Representante.Comissao = Representante.Faturado * (Representante.User.Porcentagem / 100);

                BaixaFaturamento.AddRange(BaixaDenuo.Where(x => x.TipoCliente == "S").ToList());
                BaixaFaturamento.AddRange(BaixaInter.Where(x => x.TipoCliente == "S").ToList());

            }

        }

        public IActionResult OnPostExport(string Mes, string Ano, int IdTime)
        {
            string Tempo = $"{Mes}/01/{Ano}";

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
                         select new RelatorioFaturamentoTime
                         {
                             NF = SD20.D2Doc,
                             Cliente = SA10.A1Nome,
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
                x.Cliente,
                x.Data,
                x.Login,
                x.Gestor,
                x.Linha,
                x.DOR,
                x.Codigo
            }).Select(x => new RelatorioFaturamentoTime
            {
                Empresa = "INTERMEDIC",
                NF = x.Key.NF,
                Cliente = x.Key.Cliente,
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
                Empresa = "INTERMEDIC",
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
                              join SA30 in ProtheusInter.Sa3010s on SA10.A1Vend equals SA30.A3Cod
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
                                  Gestor = SA30.A3Xlogsup,
                                  Empresa = "INTERMEDIC"
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
                          select new RelatorioFaturamentoTime
                          {
                              NF = SD20.D2Doc,
                              Cliente = SA10.A1Nome,
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
                x.Cliente,
                x.Login,
                x.Gestor,
                x.Data,
                x.Linha,
                x.Codigo
            }).Select(x => new RelatorioFaturamentoTime
            {
                Empresa = "DENUO",
                NF = x.Key.NF,
                Cliente = x.Key.Cliente,
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
                Empresa = "DENUO",
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

            var BaixaDenuo = (from SE50 in ProtheusInter.Se5010s
                              join SE10 in ProtheusInter.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                              equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                              join SA10 in ProtheusInter.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                              join SA30 in ProtheusInter.Sa3010s on SA10.A1Vend equals SA30.A3Cod
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
                                  Gestor = SA30.A3Xlogsup,
                                  Empresa = "DENUO"
                              }
                              ).ToList();

            #endregion

            #endregion

            var usuario = SGID.Times.FirstOrDefault(x => x.Id == IdTime);

            var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == usuario.Id).ToList();

            Representante = new TimeRH
            {
                User = usuario
            };

            if (Representante.User.TipoFaturamento != "S" && Representante.User.TipoFaturamento != "L")
            {
                Representante.Faturado += resultadoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total);

                Representante.Faturado += resultadoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total);

                Faturamento.AddRange(resultadoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
                Faturamento.AddRange(resultadoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
                Devolucao.AddRange(DevolucaoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
                Devolucao.AddRange(DevolucaoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());

                if (usuario.Integrante.ToUpper() == "MICHEL.SAMPAIO")
                {
                    Representante.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                    Representante.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                    GestorFaturamento.AddRange(resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").ToList());
                    GestorFaturamento.AddRange(resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").ToList());
                    GestorDevolucao.AddRange(DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").ToList());
                    GestorDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").ToList());
                }
                else
                {
                    Representante.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper() && x.DOR != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper() && x.DOR != "082").Sum(x => x.Total);
                    Representante.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper().ToUpper()).Sum(x => x.Total);
                    GestorFaturamento.AddRange(resultadoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                    GestorFaturamento.AddRange(resultadoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                    GestorDevolucao.AddRange(DevolucaoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                    GestorDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                }


                Produtos.ForEach(prod =>
                {
                    int i = 0;
                    if (usuario.Integrante.ToUpper() == "TIAGO.FONSECA")
                    {
                        if (i == 0)
                        {
                            Representante.FaturadoProduto += resultadoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total) - DevolucaoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total);
                            Representante.FaturadoProduto += resultadoInter.Where(x => (x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total) - DevolucaoInter.Where(x => (x.Linha.Trim() == prod.Produto && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total);
                            LinhaFaturamento.AddRange(resultadoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").ToList());
                            LinhaDevolucao.AddRange(DevolucaoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").ToList());
                            i++;
                        }
                        else
                        {
                            Representante.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                            Representante.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                            LinhaFaturamento.AddRange(resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").ToList());
                            LinhaDevolucao.AddRange(DevolucaoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").ToList());
                            LinhaFaturamento.AddRange(resultadoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO")).ToList());
                            LinhaDevolucao.AddRange(DevolucaoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO")).ToList());
                        }
                    }
                    else
                    {
                        Representante.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                        Representante.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                        LinhaFaturamento.AddRange(resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                        LinhaDevolucao.AddRange(DevolucaoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                        LinhaFaturamento.AddRange(resultadoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                        LinhaDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                    }
                });


                Representante.Comissao = Representante.Faturado * (Representante.User.Porcentagem / 100);
                Representante.ComissaoEquipe = Representante.FaturadoEquipe * (Representante.User.PorcentagemSeg / 100);
                Representante.ComissaoProduto = Representante.FaturadoProduto * (Representante.User.PorcentagemGenProd / 100);
            }
            else if (Representante.User.TipoFaturamento != "S")
            {


                Representante.Faturado = BaixaDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").Sum(x => x.TotalBaixado);

                Representante.Comissao = Representante.Faturado * (Representante.User.Porcentagem / 100);

                BaixaFaturamento.AddRange(BaixaDenuo.Where(x => x.CodigoCliente == "000011" || x.CodigoCliente == "000012").ToList());
            }
            else
            {

                Representante.Faturado = BaixaDenuo.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado) + BaixaInter.Where(x => x.TipoCliente == "S").Sum(x => x.TotalBaixado);

                Representante.Comissao = Representante.Faturado * (Representante.User.Porcentagem / 100);

                BaixaFaturamento.AddRange(BaixaDenuo.Where(x => x.TipoCliente == "S").ToList());
                BaixaFaturamento.AddRange(BaixaInter.Where(x => x.TipoCliente == "S").ToList());

            }


            #region Excel
            using ExcelPackage package = new ExcelPackage();
            package.Workbook.Worksheets.Add("RelatorioComissaoRH");

            var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "RelatorioComissaoRH");

            if (Representante.User.TipoFaturamento != "S" && Representante.User.TipoFaturamento != "L")
            {

                sheet.Cells[1, 1].Value = "Faturamento Venda Direta";
                sheet.Cells[1, 1, 1, 8].Merge = true;

                sheet.Cells[2, 1].Value = "Empresa";
                sheet.Cells[2, 2].Value = "NF";
                sheet.Cells[2, 3].Value = "CLIENTE";
                sheet.Cells[2, 4].Value = "VALOR";
                sheet.Cells[2, 5].Value = "COMISSÃO";
                sheet.Cells[2, 6].Value = "VENDEDOR";
                sheet.Cells[2, 7].Value = "GESTOR";
                sheet.Cells[2, 8].Value = "LINHA";

                int i = 3;

                Faturamento.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Empresa;
                    sheet.Cells[i, 2].Value = Pedido.NF;
                    sheet.Cells[i, 3].Value = Pedido.Cliente;
                    sheet.Cells[i, 4].Value = Pedido.Total;
                    sheet.Cells[i, 5].Value = Pedido.Total * (Representante.User.Porcentagem / 100);
                    sheet.Cells[i, 6].Value = Pedido.Login;
                    sheet.Cells[i, 7].Value = Pedido.Gestor;
                    sheet.Cells[i, 8].Value = Pedido.Linha;

                    i++;
                });

                i++;

                sheet.Cells[i, 1].Value = "Devolução Venda Direta";
                sheet.Cells[i, 1, i, 8].Merge = true;
                i++;

                sheet.Cells[i, 1].Value = "Empresa";
                sheet.Cells[i, 2].Value = "NF";
                sheet.Cells[i, 3].Value = "CLIENTE";
                sheet.Cells[i, 4].Value = "VALOR";
                sheet.Cells[i, 5].Value = "COMISSÃO";
                sheet.Cells[i, 6].Value = "VENDEDOR";
                sheet.Cells[i, 7].Value = "GESTOR";
                sheet.Cells[i, 8].Value = "LINHA";
                i++;        

                Devolucao.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Empresa;
                    sheet.Cells[i, 2].Value = Pedido.D1Nfori;
                    sheet.Cells[i, 3].Value = Pedido.Nome;
                    sheet.Cells[i, 4].Value = Pedido.Total;
                    sheet.Cells[i, 5].Value = Pedido.Total * (Representante.User.Porcentagem / 100);
                    sheet.Cells[i, 6].Value = Pedido.Login;
                    sheet.Cells[i, 7].Value = Pedido.Gestor;
                    sheet.Cells[i, 8].Value = Pedido.Linha;

                    i++;
                });

                i++;

                sheet.Cells[i, 1].Value = "Faturamento Sobre Equipe";
                sheet.Cells[i, 1, i, 8].Merge = true;

                i++;

                sheet.Cells[i, 1].Value = "Empresa";
                sheet.Cells[i, 2].Value = "NF";
                sheet.Cells[i, 3].Value = "CLIENTE";
                sheet.Cells[i, 4].Value = "VALOR";
                sheet.Cells[i, 5].Value = "COMISSÃO";
                sheet.Cells[i, 6].Value = "VENDEDOR";
                sheet.Cells[i, 7].Value = "GESTOR";
                sheet.Cells[i, 8].Value = "LINHA";
                i++;

                GestorFaturamento.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Empresa;
                    sheet.Cells[i, 2].Value = Pedido.NF;
                    sheet.Cells[i, 3].Value = Pedido.Cliente;
                    sheet.Cells[i, 4].Value = Pedido.Total;
                    sheet.Cells[i, 5].Value = Pedido.Total * (Representante.User.PorcentagemSeg / 100);
                    sheet.Cells[i, 6].Value = Pedido.Login;
                    sheet.Cells[i, 7].Value = Pedido.Gestor;
                    sheet.Cells[i, 8].Value = Pedido.Linha;

                    i++;
                });

                i++;

                sheet.Cells[i, 1].Value = "Devolução Sobre Equipe";
                sheet.Cells[i, 1, i, 8].Merge = true;

                i++;

                sheet.Cells[i, 1].Value = "Empresa";
                sheet.Cells[i, 2].Value = "NF";
                sheet.Cells[i, 3].Value = "CLIENTE";
                sheet.Cells[i, 4].Value = "VALOR";
                sheet.Cells[i, 5].Value = "COMISSÃO";
                sheet.Cells[i, 6].Value = "VENDEDOR";
                sheet.Cells[i, 7].Value = "GESTOR";
                sheet.Cells[i, 8].Value = "LINHA";
                i++;

                GestorDevolucao.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Empresa;
                    sheet.Cells[i, 2].Value = Pedido.D1Nfori;
                    sheet.Cells[i, 3].Value = Pedido.Nome;
                    sheet.Cells[i, 4].Value = Pedido.Total;
                    sheet.Cells[i, 5].Value = Pedido.Total * (Representante.User.PorcentagemSeg / 100);
                    sheet.Cells[i, 6].Value = Pedido.Login;
                    sheet.Cells[i, 7].Value = Pedido.Gestor;
                    sheet.Cells[i, 8].Value = Pedido.Linha;

                    i++;
                });

                i++;

                sheet.Cells[i, 1].Value = "Faturamento Sobre Linha";
                sheet.Cells[i, 1, i, 8].Merge = true;

                i++;

                sheet.Cells[i, 1].Value = "Empresa";
                sheet.Cells[i, 2].Value = "NF";
                sheet.Cells[i, 3].Value = "CLIENTE";
                sheet.Cells[i, 4].Value = "VALOR";
                sheet.Cells[i, 5].Value = "COMISSÃO";
                sheet.Cells[i, 6].Value = "VENDEDOR";
                sheet.Cells[i, 7].Value = "GESTOR";
                sheet.Cells[i, 8].Value = "LINHA";

                i++;

                LinhaFaturamento.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Empresa;
                    sheet.Cells[i, 2].Value = Pedido.NF;
                    sheet.Cells[i, 3].Value = Pedido.Cliente;
                    sheet.Cells[i, 4].Value = Pedido.Total;
                    sheet.Cells[i, 5].Value = Pedido.Total * (Representante.User.PorcentagemGenProd / 100);
                    sheet.Cells[i, 6].Value = Pedido.Login;
                    sheet.Cells[i, 7].Value = Pedido.Gestor;
                    sheet.Cells[i, 8].Value = Pedido.Linha;

                    i++;
                });

                i++;

                sheet.Cells[i, 1].Value = "Devolução Sobre Linha";
                sheet.Cells[i, 1, i, 8].Merge = true;

                i++;

                sheet.Cells[i, 1].Value = "Empresa";
                sheet.Cells[i, 2].Value = "NF";
                sheet.Cells[i, 3].Value = "CLIENTE";
                sheet.Cells[i, 4].Value = "VALOR";
                sheet.Cells[i, 5].Value = "COMISSÃO";
                sheet.Cells[i, 6].Value = "VENDEDOR";
                sheet.Cells[i, 7].Value = "GESTOR";
                sheet.Cells[i, 8].Value = "LINHA";

                i++;

                LinhaDevolucao.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Empresa;
                    sheet.Cells[i, 2].Value = Pedido.D1Nfori;
                    sheet.Cells[i, 3].Value = Pedido.Nome;
                    sheet.Cells[i, 4].Value = Pedido.Total;
                    sheet.Cells[i, 5].Value = Pedido.Total * (Representante.User.PorcentagemGenProd / 100);
                    sheet.Cells[i, 6].Value = Pedido.Login;
                    sheet.Cells[i, 7].Value = Pedido.Gestor;
                    sheet.Cells[i, 8].Value = Pedido.Linha;

                    i++;
                });
            }
            else
            {
                sheet.Cells[1, 1].Value = "Baixada no Sistema";
                sheet.Cells[1, 1, 1, 8].Merge = true;

                sheet.Cells[2, 1].Value = "Empresa";
                sheet.Cells[2, 2].Value = "NF";
                sheet.Cells[2, 3].Value = "CLIENTE";
                sheet.Cells[2, 4].Value = "VALOR";
                sheet.Cells[2, 5].Value = "COMISSÃO";
                sheet.Cells[2, 6].Value = "VENDEDOR";
                sheet.Cells[2, 7].Value = "GESTOR";
                sheet.Cells[2, 8].Value = "LINHA";

                int i = 3;

                BaixaFaturamento.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Empresa;
                    sheet.Cells[i, 2].Value = Pedido.Numero;
                    sheet.Cells[i, 3].Value = Pedido.CliFor;
                    sheet.Cells[i, 4].Value = Pedido.TotalBaixado;
                    sheet.Cells[i, 5].Value = Pedido.TotalBaixado * (Representante.User.Porcentagem / 100);
                    sheet.Cells[i, 6].Value = Pedido.Login;
                    sheet.Cells[i, 7].Value = Pedido.Gestor;
                    sheet.Cells[i, 8].Value = Pedido.Linha;

                    i++;
                });
            }


            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            #endregion

            //sheet.Protection.IsProtected = true;
            using MemoryStream stream = new MemoryStream();
            package.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"RelatorioComissoesRH{Mes}{Ano}{usuario.Integrante.ToUpper().Replace(".","")}.xlsx");

        }
    }
}
