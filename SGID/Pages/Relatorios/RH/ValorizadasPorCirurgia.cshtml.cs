using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.AdmVendas;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.RH
{
    public class ValorizadasPorCirurgiaModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }

        private TOTVSINTERContext ProtheusInter { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioCirurgiasValorizadas> Relatorio = new List<RelatorioCirurgiasValorizadas>();
        public ValorizadasPorCirurgiaModel(TOTVSINTERContext inter, TOTVSDENUOContext denuo, ApplicationDbContext sgid)
        {
            ProtheusInter = inter;
            ProtheusDenuo = denuo;
            SGID = sgid;
        }
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public List<string> Pedidos { get; set; } = new List<string>();
        public List<string> Clientes { get; set; } = new List<string>();
        public List<string> Grupos { get; set; } = new List<string>();

        public string Cliente { get; set; }
        public string Grupo { get; set; }

        public void OnGet(string Empresa)
        {
           
        }

        public IActionResult OnPostAsync(DateTime Inicio,DateTime Fim, string Empresa)
        {
            try
            {
                string DataInicio = Inicio.ToString("yyyy/MM/dd").Replace("/","");

                string DataFim = Fim.ToString("yyyy/MM/dd").Replace("/", "");

                this.Inicio = Inicio;
                this.Fim = Fim;

                var user = User.Identity.Name.Split("@")[0].ToUpper();

                if (Empresa == "01")
                {
                    var RelatorioInter = (from SC50 in ProtheusInter.Sc5010s
                                          join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                          join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                          join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                          join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                          join PA10 in ProtheusInter.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                          from c in Sr.DefaultIfEmpty()
                                          join SUA10 in ProtheusInter.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                          from a in Rs.DefaultIfEmpty()
                                          where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                          && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                          && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                          select new RelatorioCirurgiasValorizadas
                                          {
                                              DTCirurgia = SC50.C5XDtcir,
                                              Filial = SC50.C5Filial,
                                              NumPedido = SC50.C5Num,
                                              Agendamento = a.UaUnumage,
                                              TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC50.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                                              Vendedor = SC50.C5Nomvend,
                                              Medico = SC50.C5XNmmed,
                                              Paciente = SC50.C5XNmpac,
                                              ClienteEnt = SC50.C5XNmpla,
                                              QTDPedido = SC60.C6Qtdven,
                                              QTDFaturada = SC60.C6Qtdent,
                                              VLUnitario = SC60.C6Prcven,
                                              TotalPedido = SC60.C6Valor,
                                              TotalFat = SC60.C6Qtdent * SC60.C6Prcven,
                                              Faturado = SC50.C5Nota != "" || (SC50.C5Liberok == "E" && SC50.C5Blq == "") ? "S" : "N"

                                          }).GroupBy(x => new
                                          {
                                              x.DTCirurgia,
                                              x.Filial,
                                              x.NumPedido,
                                              x.Agendamento,
                                              x.TipoOper,
                                              x.Vendedor,
                                              x.Medico,
                                              x.Paciente,
                                              x.ClienteEnt,
                                              x.Faturado
                                          });

                    Relatorio = RelatorioInter.Select(x => new RelatorioCirurgiasValorizadas
                    {
                        DTCirurgia = $"{x.Key.DTCirurgia.Substring(6, 2)}/{x.Key.DTCirurgia.Substring(4, 2)}/{x.Key.DTCirurgia.Substring(0, 4)}",
                        Filial = x.Key.Filial,
                        NumPedido = x.Key.NumPedido,
                        Agendamento = x.Key.Agendamento,
                        TipoOper = x.Key.TipoOper,
                        Vendedor = x.Key.Vendedor,
                        Medico = x.Key.Medico,
                        Paciente = x.Key.Paciente,
                        ClienteEnt = x.Key.ClienteEnt,
                        QTDPedido = x.Sum(c => c.QTDPedido),
                        QTDFaturada = x.Sum(c => c.QTDFaturada),
                        TotalPedido = x.Sum(c => c.TotalPedido),
                        TotalFat = x.Sum(c => c.TotalFat),
                        Faturado = x.Key.Faturado
                    }).OrderBy(x => x.Vendedor).ToList();
                }
                else
                {

                    if (user != "TIAGO.FONSECA" && !User.IsInRole("Admin"))
                    {
                        if (user == "ARTEMIO.COSTA") user = "LEONARDO.BRITO";
                        #region Gestor

                        var RelatorioDenuo = (from SC50 in ProtheusDenuo.Sc5010s
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
                                              && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                              && (SA30.A3Xlogin == user || SA30.A3Xlogsup == user)
                                              select new RelatorioCirurgiasValorizadas
                                              {
                                                  DTCirurgia = SC50.C5XDtcir,
                                                  Filial = SC50.C5Filial,
                                                  NumPedido = SC50.C5Num,
                                                  Agendamento = a.UaUnumage,
                                                  TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC50.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                                                  Vendedor = SC50.C5Nomvend,
                                                  Medico = SC50.C5XNmmed,
                                                  Paciente = SC50.C5XNmpac,
                                                  ClienteEnt = SC50.C5XNmpla,
                                                  QTDPedido = SC60.C6Qtdven,
                                                  QTDFaturada = SC60.C6Qtdent,
                                                  VLUnitario = SC60.C6Prcven,
                                                  TotalPedido = SC60.C6Valor,
                                                  TotalFat = SC60.C6Qtdent * SC60.C6Prcven,
                                                  Faturado = SC50.C5Nota != "" || (SC50.C5Liberok == "E" && SC50.C5Blq == "") ? "S" : "N"

                                              }).GroupBy(x => new
                                              {
                                                  x.DTCirurgia,
                                                  x.Filial,
                                                  x.NumPedido,
                                                  x.Agendamento,
                                                  x.TipoOper,
                                                  x.Vendedor,
                                                  x.Medico,
                                                  x.Paciente,
                                                  x.ClienteEnt,
                                                  x.Faturado
                                              });


                        Relatorio = RelatorioDenuo.Select(x => new RelatorioCirurgiasValorizadas
                        {
                            DTCirurgia = $"{x.Key.DTCirurgia.Substring(6, 2)}/{x.Key.DTCirurgia.Substring(4, 2)}/{x.Key.DTCirurgia.Substring(0, 4)}",
                            Filial = x.Key.Filial,
                            NumPedido = x.Key.NumPedido,
                            Agendamento = x.Key.Agendamento,
                            TipoOper = x.Key.TipoOper,
                            Vendedor = x.Key.Vendedor,
                            Medico = x.Key.Medico,
                            Paciente = x.Key.Paciente,
                            ClienteEnt = x.Key.ClienteEnt,
                            QTDPedido = x.Sum(c => c.QTDPedido),
                            QTDFaturada = x.Sum(c => c.QTDFaturada),
                            TotalPedido = x.Sum(c => c.TotalPedido),
                            TotalFat = x.Sum(c => c.TotalFat),
                            Faturado = x.Key.Faturado
                        }).OrderBy(x => x.Vendedor).ToList();


                        #endregion
                    }
                    else
                    {
                        #region Tiago

                        var RelatorioDenuo = (from SC50 in ProtheusDenuo.Sc5010s
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
                                              && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                              select new RelatorioCirurgiasValorizadas
                                              {
                                                  DTCirurgia = SC50.C5XDtcir,
                                                  Filial = SC50.C5Filial,
                                                  NumPedido = SC50.C5Num,
                                                  Agendamento = a.UaUnumage,
                                                  TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC50.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                                                  Vendedor = SC50.C5Nomvend,
                                                  Medico = SC50.C5XNmmed,
                                                  Paciente = SC50.C5XNmpac,
                                                  ClienteEnt = SC50.C5XNmpla,
                                                  QTDPedido = SC60.C6Qtdven,
                                                  QTDFaturada = SC60.C6Qtdent,
                                                  VLUnitario = SC60.C6Prcven,
                                                  TotalPedido = SC60.C6Valor,
                                                  TotalFat = SC60.C6Qtdent * SC60.C6Prcven,
                                                  Faturado = SC50.C5Nota != "" || (SC50.C5Liberok == "E" && SC50.C5Blq == "") ? "S" : "N"

                                              }).GroupBy(x => new
                                              {
                                                  x.DTCirurgia,
                                                  x.Filial,
                                                  x.NumPedido,
                                                  x.Agendamento,
                                                  x.TipoOper,
                                                  x.Vendedor,
                                                  x.Medico,
                                                  x.Paciente,
                                                  x.ClienteEnt,
                                                  x.Faturado
                                              });


                        Relatorio = RelatorioDenuo.Select(x => new RelatorioCirurgiasValorizadas
                        {
                            DTCirurgia = $"{x.Key.DTCirurgia.Substring(6, 2)}/{x.Key.DTCirurgia.Substring(4, 2)}/{x.Key.DTCirurgia.Substring(0, 4)}",
                            Filial = x.Key.Filial,
                            NumPedido = x.Key.NumPedido,
                            Agendamento = x.Key.Agendamento,
                            TipoOper = x.Key.TipoOper,
                            Vendedor = x.Key.Vendedor,
                            Medico = x.Key.Medico,
                            Paciente = x.Key.Paciente,
                            ClienteEnt = x.Key.ClienteEnt,
                            QTDPedido = x.Sum(c => c.QTDPedido),
                            QTDFaturada = x.Sum(c => c.QTDFaturada),
                            TotalPedido = x.Sum(c => c.TotalPedido),
                            TotalFat = x.Sum(c => c.TotalFat),
                            Faturado = x.Key.Faturado
                        }).OrderBy(x => x.Vendedor).ToList();

                        #endregion
                    }


                }

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ValorizadasPorCirurgia", user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime Inicio, DateTime Fim, string Empresa)
        {
            try
            {
                string DataInicio = Inicio.ToString("yyyy/MM/dd").Replace("/", "");

                string DataFim = Fim.ToString("yyyy/MM/dd").Replace("/", "");

                var user = User.Identity.Name.Split("@")[0].ToUpper();

                if (Empresa == "01")
                {
                    var RelatorioInter = (from SC50 in ProtheusInter.Sc5010s
                                          join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                          join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                          join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                          join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                          join PA10 in ProtheusInter.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                          from c in Sr.DefaultIfEmpty()
                                          join SUA10 in ProtheusInter.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                          from a in Rs.DefaultIfEmpty()
                                          where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                          && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim
                                          && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                          select new RelatorioCirurgiasValorizadas
                                          {
                                              DTCirurgia = SC50.C5XDtcir,
                                              Filial = SC50.C5Filial,
                                              NumPedido = SC50.C5Num,
                                              Agendamento = a.UaUnumage,
                                              TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC50.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                                              Vendedor = SC50.C5Nomvend,
                                              Medico = SC50.C5XNmmed,
                                              Paciente = SC50.C5XNmpac,
                                              ClienteEnt = SC50.C5XNmpla,
                                              QTDPedido = SC60.C6Qtdven,
                                              QTDFaturada = SC60.C6Qtdent,
                                              VLUnitario = SC60.C6Prcven,
                                              TotalPedido = SC60.C6Valor,
                                              TotalFat = SC60.C6Qtdent * SC60.C6Prcven,
                                              Faturado = SC50.C5Nota != "" || (SC50.C5Liberok == "E" && SC50.C5Blq == "") ? "S" : "N"

                                          }).GroupBy(x => new
                                          {
                                              x.DTCirurgia,
                                              x.Filial,
                                              x.NumPedido,
                                              x.Agendamento,
                                              x.TipoOper,
                                              x.Vendedor,
                                              x.Medico,
                                              x.Paciente,
                                              x.ClienteEnt,
                                              x.Faturado
                                          });

                    Relatorio = RelatorioInter.Select(x => new RelatorioCirurgiasValorizadas
                    {
                        DTCirurgia = $"{x.Key.DTCirurgia.Substring(6, 2)}/{x.Key.DTCirurgia.Substring(4, 2)}/{x.Key.DTCirurgia.Substring(0, 4)}",
                        Filial = x.Key.Filial,
                        NumPedido = x.Key.NumPedido,
                        Agendamento = x.Key.Agendamento,
                        TipoOper = x.Key.TipoOper,
                        Vendedor = x.Key.Vendedor,
                        Medico = x.Key.Medico,
                        Paciente = x.Key.Paciente,
                        ClienteEnt = x.Key.ClienteEnt,
                        QTDPedido = x.Sum(c => c.QTDPedido),
                        QTDFaturada = x.Sum(c => c.QTDFaturada),
                        TotalPedido = x.Sum(c => c.TotalPedido),
                        TotalFat = x.Sum(c => c.TotalFat),
                        Faturado = x.Key.Faturado
                    }).OrderBy(x => x.Vendedor).ToList();
                }
                else
                {

                    if (user != "TIAGO.FONSECA" && !User.IsInRole("Admin"))
                    {
                        if (user == "ARTEMIO.COSTA") user = "LEONARDO.BRITO";
                        #region Gestor

                        var RelatorioDenuo = (from SC50 in ProtheusDenuo.Sc5010s
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
                                              && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                              && (SA30.A3Xlogin == user || SA30.A3Xlogsup == user)
                                              select new RelatorioCirurgiasValorizadas
                                              {
                                                  DTCirurgia = SC50.C5XDtcir,
                                                  Filial = SC50.C5Filial,
                                                  NumPedido = SC50.C5Num,
                                                  Agendamento = a.UaUnumage,
                                                  TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC50.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                                                  Vendedor = SC50.C5Nomvend,
                                                  Medico = SC50.C5XNmmed,
                                                  Paciente = SC50.C5XNmpac,
                                                  ClienteEnt = SC50.C5XNmpla,
                                                  QTDPedido = SC60.C6Qtdven,
                                                  QTDFaturada = SC60.C6Qtdent,
                                                  VLUnitario = SC60.C6Prcven,
                                                  TotalPedido = SC60.C6Valor,
                                                  TotalFat = SC60.C6Qtdent * SC60.C6Prcven,
                                                  Faturado = SC50.C5Nota != "" || (SC50.C5Liberok == "E" && SC50.C5Blq == "") ? "S" : "N"

                                              }).GroupBy(x => new
                                              {
                                                  x.DTCirurgia,
                                                  x.Filial,
                                                  x.NumPedido,
                                                  x.Agendamento,
                                                  x.TipoOper,
                                                  x.Vendedor,
                                                  x.Medico,
                                                  x.Paciente,
                                                  x.ClienteEnt,
                                                  x.Faturado
                                              });


                        Relatorio = RelatorioDenuo.Select(x => new RelatorioCirurgiasValorizadas
                        {
                            DTCirurgia = $"{x.Key.DTCirurgia.Substring(6, 2)}/{x.Key.DTCirurgia.Substring(4, 2)}/{x.Key.DTCirurgia.Substring(0, 4)}",
                            Filial = x.Key.Filial,
                            NumPedido = x.Key.NumPedido,
                            Agendamento = x.Key.Agendamento,
                            TipoOper = x.Key.TipoOper,
                            Vendedor = x.Key.Vendedor,
                            Medico = x.Key.Medico,
                            Paciente = x.Key.Paciente,
                            ClienteEnt = x.Key.ClienteEnt,
                            QTDPedido = x.Sum(c => c.QTDPedido),
                            QTDFaturada = x.Sum(c => c.QTDFaturada),
                            TotalPedido = x.Sum(c => c.TotalPedido),
                            TotalFat = x.Sum(c => c.TotalFat),
                            Faturado = x.Key.Faturado
                        }).OrderBy(x => x.Vendedor).ToList();

                       
                        #endregion
                    }
                    else
                    {
                        #region Tiago

                        var RelatorioDenuo = (from SC50 in ProtheusDenuo.Sc5010s
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
                                              && (SC50.C5Utpoper == "F" || SC50.C5Utpoper == "K") && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                              select new RelatorioCirurgiasValorizadas
                                              {
                                                  DTCirurgia = SC50.C5XDtcir,
                                                  Filial = SC50.C5Filial,
                                                  NumPedido = SC50.C5Num,
                                                  Agendamento = a.UaUnumage,
                                                  TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC50.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                                                  Vendedor = SC50.C5Nomvend,
                                                  Medico = SC50.C5XNmmed,
                                                  Paciente = SC50.C5XNmpac,
                                                  ClienteEnt = SC50.C5XNmpla,
                                                  QTDPedido = SC60.C6Qtdven,
                                                  QTDFaturada = SC60.C6Qtdent,
                                                  VLUnitario = SC60.C6Prcven,
                                                  TotalPedido = SC60.C6Valor,
                                                  TotalFat = SC60.C6Qtdent * SC60.C6Prcven,
                                                  Faturado = SC50.C5Nota != "" || (SC50.C5Liberok == "E" && SC50.C5Blq == "") ? "S" : "N"

                                              }).GroupBy(x => new
                                              {
                                                  x.DTCirurgia,
                                                  x.Filial,
                                                  x.NumPedido,
                                                  x.Agendamento,
                                                  x.TipoOper,
                                                  x.Vendedor,
                                                  x.Medico,
                                                  x.Paciente,
                                                  x.ClienteEnt,
                                                  x.Faturado
                                              });


                        Relatorio = RelatorioDenuo.Select(x => new RelatorioCirurgiasValorizadas
                        {
                            DTCirurgia = $"{x.Key.DTCirurgia.Substring(6, 2)}/{x.Key.DTCirurgia.Substring(4, 2)}/{x.Key.DTCirurgia.Substring(0, 4)}",
                            Filial = x.Key.Filial,
                            NumPedido = x.Key.NumPedido,
                            Agendamento = x.Key.Agendamento,
                            TipoOper = x.Key.TipoOper,
                            Vendedor = x.Key.Vendedor,
                            Medico = x.Key.Medico,
                            Paciente = x.Key.Paciente,
                            ClienteEnt = x.Key.ClienteEnt,
                            QTDPedido = x.Sum(c => c.QTDPedido),
                            QTDFaturada = x.Sum(c => c.QTDFaturada),
                            TotalPedido = x.Sum(c => c.TotalPedido),
                            TotalFat = x.Sum(c => c.TotalFat),
                            Faturado = x.Key.Faturado
                        }).OrderBy(x => x.Vendedor).ToList();

                        #endregion
                    }

                    
                }

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Cirurgias Valorizadas");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Cirurgias Valorizadas");

                sheet.Cells[1, 1].Value = "Data Cirurgia";
                sheet.Cells[1, 2].Value = "Filial";
                sheet.Cells[1, 3].Value = "Num. Pedido";
                sheet.Cells[1, 4].Value = "Agendamento";
                sheet.Cells[1, 5].Value = "Tipo de Operação";
                sheet.Cells[1, 6].Value = "Vendedor";
                sheet.Cells[1, 7].Value = "Médico";
                sheet.Cells[1, 8].Value = "Paciente";
                sheet.Cells[1, 9].Value = "Cliente Entrega";
                sheet.Cells[1, 10].Value = "Convênio";
                sheet.Cells[1, 11].Value = "QTD Pedido";
                sheet.Cells[1, 12].Value = "QTD Faturada";
                sheet.Cells[1, 13].Value = "Total Pedido";
                sheet.Cells[1, 14].Value = "Total Fat.";
                sheet.Cells[1, 15].Value = "Faturado ?";

                int i = 2;
                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.DTCirurgia;
                    sheet.Cells[i, 2].Value = Pedido.Filial;
                    sheet.Cells[i, 3].Value = Pedido.NumPedido;
                    sheet.Cells[i, 4].Value = Pedido.Agendamento;
                    sheet.Cells[i, 5].Value = Pedido.TipoOper;
                    sheet.Cells[i, 6].Value = Pedido.Vendedor;
                    sheet.Cells[i, 7].Value = Pedido.Medico;
                    sheet.Cells[i, 8].Value = Pedido.Paciente;
                    sheet.Cells[i, 9].Value = Pedido.ClienteEnt;
                    sheet.Cells[i, 10].Value = Pedido.Convenio;
                    sheet.Cells[i, 11].Value = Pedido.QTDPedido;
                    sheet.Cells[i, 12].Value = Pedido.QTDFaturada;
                    sheet.Cells[i, 13].Value = Pedido.TotalPedido;
                    sheet.Cells[i, 14].Value = Pedido.TotalFat;
                    sheet.Cells[i, 15].Value = Pedido.Faturado;

                    i++;
                });

                sheet.Cells[i, 1].Value = "Total Geral";
                sheet.Cells[i, 11].Value = Relatorio.Sum(x => x.QTDPedido);
                sheet.Cells[i, 12].Value = Relatorio.Sum(x => x.QTDFaturada);
                sheet.Cells[i, 13].Value = Relatorio.Sum(x => x.TotalPedido);
                sheet.Cells[i, 14].Value = Relatorio.Sum(x => x.TotalFat);

                sheet.Cells[i, 1, i, 10].Merge = true;

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CirurgiasValorizadasDenuo.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasValorizadasDiretoria Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
