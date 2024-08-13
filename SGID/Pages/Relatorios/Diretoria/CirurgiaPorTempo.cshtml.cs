using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.AdmVendas;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.Diretoria
{
    public class CirurgiaPorTempoModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public List<RelatorioCirurgiaAFaturar> Relatorio { get; set; } = new List<RelatorioCirurgiaAFaturar>();
        public List<RelatorioCirurgiaAFaturar> RelatorioInter { get; set; } = new List<RelatorioCirurgiaAFaturar>();
        public List<string> Pedidos { get; set; } = new List<string>();
        public List<string> Clientes { get; set; } = new List<string>();

        public List<string> Grupos { get; set; } = new List<string>();
        public double Total { get; set; }
        public string Anging { get; set; }
        public string Cliente { get; set; }
        public string Empresa { get; set; }
        public string Grupo { get; set; }

        public CirurgiaPorTempoModel(TOTVSDENUOContext context, ApplicationDbContext sgid,TOTVSINTERContext inter)
        {
            Protheus = context;
            SGID = sgid;
            ProtheusInter = inter;
        }

        public void OnGet(string id)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();

                #region Relatorio
                Relatorio = (from SC5 in Protheus.Sc5010s
                             join SC6 in Protheus.Sc6010s on new { Filial = SC5.C5Filial, Num = SC5.C5Num } equals new { Filial = SC6.C6Filial, Num = SC6.C6Num }
                             join SA1 in Protheus.Sa1010s on SC5.C5Cliente equals SA1.A1Cod
                             join SA3 in Protheus.Sa3010s on SC5.C5Vend1 equals SA3.A3Cod
                             join SF4 in Protheus.Sf4010s on SC6.C6Tes equals SF4.F4Codigo
                             join SB1 in Protheus.Sb1010s on SC6.C6Produto equals SB1.B1Cod
                             join SUA in Protheus.Sua010s on SC5.C5Uproces equals SUA.UaNum into Sr
                             from c in Sr.DefaultIfEmpty()
                             join SX5 in Protheus.Sx5010s on new { Grupo = SA1.A1Xgrinte, Filial = SC5.C5Filial, Tabela = "Z3" } equals new { Grupo = SX5.X5Chave, Filial = SX5.X5Filial, Tabela = SX5.X5Tabela } into Se
                             from a in Se.DefaultIfEmpty()
                             where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num && a.DELET!="*"
                             && SC6.C6Nota == ""
                             && SC5.C5Nota == ""
                             && SC6.C6Blq != "R"
                             && SC6.DELET != "*"
                             && SF4.F4Codigo == SC6.C6Tes
                             && SF4.F4Duplic == "S"
                             && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                             && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                             && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T")
                             && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                             && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                             orderby SC5.C5Num, SC5.C5Emissao descending
                             select new RelatorioCirurgiaAFaturar
                             {
                                 Vendedor = SC5.C5Nomvend,
                                 Cliente = SC5.C5Nomcli,
                                 GrupoCliente = a.X5Descri,
                                 ClienteEntrega = SC5.C5Nomclie,
                                 Medico = SC5.C5XNmmed,
                                 Convenio = SC5.C5XNmpla,
                                 Cirurgia = Convert.ToInt32(SC5.C5XDtcir),
                                 Hoje = "",
                                 Dias = 0,
                                 Anging = "",
                                 Paciente = SC5.C5XNmpac,
                                 Matric = "",
                                 INPART = "",
                                 Valor = SC6.C6Valor,
                                 PVFaturamento = SC5.C5Num,
                                 Status = c.UaXstatus,
                                 DataEnvRA = " / / ",
                                 DataRecRA = " / / ",
                                 DataValorizacao = SC5.C5Xdtval,
                                 Emissao = Convert.ToInt32(SC5.C5Emissao)
                             }
                        ).GroupBy(x => new
                        {
                            x.Vendedor,
                            x.Cliente,
                            x.GrupoCliente,
                            x.ClienteEntrega,
                            x.Medico,
                            x.Convenio,
                            x.Cirurgia,
                            x.Hoje,
                            x.Dias,
                            x.Anging,
                            x.Paciente,
                            x.Matric,
                            x.INPART,
                            x.PVFaturamento,
                            x.Status,
                            x.DataEnvRA,
                            x.DataRecRA,
                            x.DataValorizacao,
                            x.Emissao
                        })
                        .Select(x => new RelatorioCirurgiaAFaturar
                        {
                            Vendedor = x.Key.Vendedor,
                            Cliente = x.Key.Cliente,
                            GrupoCliente = x.Key.GrupoCliente,
                            ClienteEntrega = x.Key.ClienteEntrega,
                            Medico = x.Key.Medico,
                            Convenio = x.Key.Convenio,
                            Cirurgia = x.Key.Cirurgia,
                            Hoje = x.Key.Hoje,
                            Dias = x.Key.Dias,
                            Anging = x.Key.Anging,
                            Paciente = x.Key.Paciente,
                            Matric = x.Key.Matric,
                            INPART = x.Key.INPART,
                            Valor = x.Sum(c => c.Valor),
                            PVFaturamento = x.Key.PVFaturamento,
                            Status = x.Key.Status,
                            DataEnvRA = x.Key.DataEnvRA,
                            DataRecRA = x.Key.DataRecRA,
                            DataValorizacao = x.Key.DataValorizacao,
                            Empresa = "DENUO",
                            Emissao = x.Key.Emissao
                        }).ToList();

                RelatorioInter = (from SC5 in ProtheusInter.Sc5010s
                                  join SC6 in ProtheusInter.Sc6010s on new { Filial = SC5.C5Filial, Num = SC5.C5Num } equals new { Filial = SC6.C6Filial, Num = SC6.C6Num }
                                  join SA1 in ProtheusInter.Sa1010s on SC5.C5Cliente equals SA1.A1Cod
                                  join SA3 in ProtheusInter.Sa3010s on SC5.C5Vend1 equals SA3.A3Cod
                                  join SF4 in ProtheusInter.Sf4010s on SC6.C6Tes equals SF4.F4Codigo
                                  join SB1 in ProtheusInter.Sb1010s on SC6.C6Produto equals SB1.B1Cod
                                  join SUA in ProtheusInter.Sua010s on SC5.C5Uproces equals SUA.UaNum into Sr
                                  from c in Sr.DefaultIfEmpty()
                                  join SX5 in ProtheusInter.Sx5010s on new { Grupo = SA1.A1Xgrinte, Tabela = "Z3" } equals new { Grupo = SX5.X5Chave, Tabela = SX5.X5Tabela } into Se
                                  from a in Se.DefaultIfEmpty()
                                  where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num && a.DELET != "*"
                                  && SC6.C6Nota == ""
                                  && SC5.C5Nota == ""
                                  && SC6.C6Blq != "R"
                                  && SC6.DELET != "*"
                                  && SF4.F4Codigo == SC6.C6Tes
                                  && SF4.F4Duplic == "S"
                                  && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                                  && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                                  && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T")
                                  && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                                  && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                                  orderby SC5.C5Num, SC5.C5Emissao descending
                                  select new RelatorioCirurgiaAFaturar
                                  {
                                      Vendedor = SC5.C5Nomvend,
                                      Cliente = SC5.C5Nomcli,
                                      GrupoCliente = a.X5Descri,
                                      ClienteEntrega = SC5.C5Nomclie,
                                      Medico = SC5.C5XNmmed,
                                      Convenio = SC5.C5XNmpla,
                                      Cirurgia = Convert.ToInt32(SC5.C5XDtcir),
                                      Hoje = "",
                                      Dias = 0,
                                      Anging = "",
                                      Paciente = SC5.C5XNmpac,
                                      Matric = "",
                                      INPART = "",
                                      Valor = SC6.C6Valor,
                                      PVFaturamento = SC5.C5Num,
                                      Status = c.UaXstatus,
                                      DataEnvRA = " / / ",
                                      DataRecRA = " / / ",
                                      DataValorizacao = SC5.C5Xdtval,
                                      Emissao = Convert.ToInt32(SC5.C5Emissao)
                                  }
                        ).GroupBy(x => new
                        {
                            x.Vendedor,
                            x.Cliente,
                            x.GrupoCliente,
                            x.ClienteEntrega,
                            x.Medico,
                            x.Convenio,
                            x.Cirurgia,
                            x.Hoje,
                            x.Dias,
                            x.Anging,
                            x.Paciente,
                            x.Matric,
                            x.INPART,
                            x.PVFaturamento,
                            x.Status,
                            x.DataEnvRA,
                            x.DataRecRA,
                            x.DataValorizacao,
                            x.Emissao
                        })
                        .Select(x => new RelatorioCirurgiaAFaturar
                        {
                            Vendedor = x.Key.Vendedor,
                            Cliente = x.Key.Cliente,
                            GrupoCliente = x.Key.GrupoCliente,
                            ClienteEntrega = x.Key.ClienteEntrega,
                            Medico = x.Key.Medico,
                            Convenio = x.Key.Convenio,
                            Cirurgia = x.Key.Cirurgia,
                            Hoje = x.Key.Hoje,
                            Dias = x.Key.Dias,
                            Anging = x.Key.Anging,
                            Paciente = x.Key.Paciente,
                            Matric = x.Key.Matric,
                            INPART = x.Key.INPART,
                            Valor = x.Sum(c => c.Valor),
                            PVFaturamento = x.Key.PVFaturamento,
                            Status = x.Key.Status,
                            DataEnvRA = x.Key.DataEnvRA,
                            DataRecRA = x.Key.DataRecRA,
                            DataValorizacao = x.Key.DataValorizacao,
                            Empresa = "INTERMEDIC",
                            Emissao = x.Key.Emissao
                        }).ToList();
                #endregion

                Relatorio = Relatorio.Concat(RelatorioInter).ToList();

                Relatorio = Relatorio.Where(x => x.Status != "C").ToList();

                if (id == "1")
                {
                    DateTime ago = DateTime.Now.AddMonths(-3);

                    var Data = Convert.ToInt32(ago.ToString("yyyy/MM/dd").Replace("/", ""));

                    Relatorio = Relatorio.Where(x => x.Cirurgia >= Data).ToList();
                }
                else if (id == "2")
                {
                    DateTime ago = DateTime.Now.AddMonths(-9);
                    DateTime After = ago.AddMonths(6);

                    var Data = Convert.ToInt32(ago.ToString("yyyy/MM/dd").Replace("/", ""));
                    var Data2 = Convert.ToInt32(After.ToString("yyyy/MM/dd").Replace("/", ""));

                    Relatorio = Relatorio.Where(x => x.Cirurgia >= Data && x.Cirurgia < Data2).ToList();
                }
                else
                {
                    DateTime ago = DateTime.Now.AddMonths(-9);

                    var Data = Convert.ToInt32(ago.ToString("yyyy/MM/dd").Replace("/", ""));

                    Relatorio = Relatorio.Where(x => x.Cirurgia < Data).ToList();
                }

                Grupos = Relatorio.Where(x=> x.GrupoCliente !="" && x.GrupoCliente!=null).Select(x => x.GrupoCliente).Distinct().ToList();
                Clientes = Relatorio.Select(x => x.Cliente).Distinct().ToList();

                var data = DateTime.Now;

                Relatorio.ForEach(x =>
                {
                    x.Hoje = data.ToString("dd/MM/yyyy");

                    var check = x.Cirurgia;

                    if (check != 0)
                    {
                        x.Dias = (int)(data - Convert.ToDateTime($"{x.Cirurgia.ToString().Substring(4, 2)}/{x.Cirurgia.ToString().Substring(6, 2)}/{x.Cirurgia.ToString().Substring(0, 4)}")).TotalDays;
                    }

                    var check2 = int.TryParse(x.DataValorizacao, out int result2);

                    if (check2)
                    {
                        x.DataValorizacao = $"{x.DataValorizacao.Substring(6, 2)}/{x.DataValorizacao.Substring(4, 2)}/{x.DataValorizacao.Substring(0, 4)}";
                    }


                    if (x.Dias <= 30)
                    {
                        x.Anging = "ATE 30";
                    }
                    else if (x.Dias <= 60)
                    {
                        x.Anging = "31 A 60";
                    }
                    else if (x.Dias <= 90)
                    {
                        x.Anging = "61 A 90";
                    }
                    else if (x.Dias <= 120)
                    {
                        x.Anging = "91 A 120";
                    }
                    else
                    {
                        x.Anging = "MAIS DE 120";
                    }
                    Total += x.Valor;
                });
                
                Relatorio = Relatorio.OrderBy(x => x.Dias).ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiaPorTempo", user);
            }
        }

        public IActionResult OnPost(string id, string NReduz, string Grupo, string Empresa)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();

                #region Relatorio
                Relatorio = (from SC5 in Protheus.Sc5010s
                             join SC6 in Protheus.Sc6010s on new { Filial = SC5.C5Filial, Num = SC5.C5Num } equals new { Filial = SC6.C6Filial, Num = SC6.C6Num }
                             join SA1 in Protheus.Sa1010s on SC5.C5Cliente equals SA1.A1Cod
                             join SA3 in Protheus.Sa3010s on SC5.C5Vend1 equals SA3.A3Cod
                             join SF4 in Protheus.Sf4010s on SC6.C6Tes equals SF4.F4Codigo
                             join SB1 in Protheus.Sb1010s on SC6.C6Produto equals SB1.B1Cod
                             join SUA in Protheus.Sua010s on SC5.C5Uproces equals SUA.UaNum into Sr
                             from c in Sr.DefaultIfEmpty()
                             join SX5 in Protheus.Sx5010s on new { Grupo = SA1.A1Xgrinte, Filial = SC5.C5Filial, Tabela = "Z3" } equals new { Grupo = SX5.X5Chave, Filial = SX5.X5Filial, Tabela = SX5.X5Tabela } into Se
                             from a in Se.DefaultIfEmpty()
                             where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num && a.DELET != "*"
                             && SC6.C6Nota == ""
                             && SC5.C5Nota == ""
                             && SC6.C6Blq != "R"
                             && SC6.DELET != "*"
                             && SF4.F4Codigo == SC6.C6Tes
                             && SF4.F4Duplic == "S"
                             && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                             && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                             && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T")
                             && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                             && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                             orderby SC5.C5Num, SC5.C5Emissao descending
                             select new RelatorioCirurgiaAFaturar
                             {
                                 Vendedor = SC5.C5Nomvend,
                                 Cliente =  SC5.C5Nomcli,
                                 GrupoCliente = a.X5Descri,
                                 ClienteEntrega = SC5.C5Nomclie,
                                 Medico = SC5.C5XNmmed,
                                 Convenio = SC5.C5XNmpla,
                                 Cirurgia = Convert.ToInt32(SC5.C5XDtcir),
                                 Hoje = "",
                                 Dias = 0,
                                 Anging = "",
                                 Paciente = SC5.C5XNmpac,
                                 Matric = "",
                                 INPART = "",
                                 Valor = SC6.C6Valor,
                                 PVFaturamento = SC5.C5Num,
                                 Status = c.UaXstatus,
                                 DataEnvRA = " / / ",
                                 DataRecRA = " / / ",
                                 DataValorizacao = SC5.C5Xdtval,
                                 Emissao = Convert.ToInt32(SC5.C5Emissao)
                             }
                        ).GroupBy(x => new
                        {
                            x.Vendedor,
                            x.Cliente,
                            x.GrupoCliente,
                            x.ClienteEntrega,
                            x.Medico,
                            x.Convenio,
                            x.Cirurgia,
                            x.Hoje,
                            x.Dias,
                            x.Anging,
                            x.Paciente,
                            x.Matric,
                            x.INPART,
                            x.PVFaturamento,
                            x.Status,
                            x.DataEnvRA,
                            x.DataRecRA,
                            x.DataValorizacao,
                            x.Emissao
                        })
                        .Select(x => new RelatorioCirurgiaAFaturar
                        {
                            Vendedor = x.Key.Vendedor,
                            Cliente = x.Key.Cliente,
                            GrupoCliente = x.Key.GrupoCliente,
                            ClienteEntrega = x.Key.ClienteEntrega,
                            Medico = x.Key.Medico,
                            Convenio = x.Key.Convenio,
                            Cirurgia = x.Key.Cirurgia,
                            Hoje = x.Key.Hoje,
                            Dias = x.Key.Dias,
                            Anging = x.Key.Anging,
                            Paciente = x.Key.Paciente,
                            Matric = x.Key.Matric,
                            INPART = x.Key.INPART,
                            Valor = x.Sum(c => c.Valor),
                            PVFaturamento = x.Key.PVFaturamento,
                            Status = x.Key.Status,
                            DataEnvRA = x.Key.DataEnvRA,
                            DataRecRA = x.Key.DataRecRA,
                            DataValorizacao = x.Key.DataValorizacao,
                            Empresa = "DENUO",
                            Emissao = x.Key.Emissao
                        }).ToList();

                RelatorioInter = (from SC5 in ProtheusInter.Sc5010s
                                  join SC6 in ProtheusInter.Sc6010s on new { Filial = SC5.C5Filial, Num = SC5.C5Num } equals new { Filial = SC6.C6Filial, Num = SC6.C6Num }
                                  join SA1 in ProtheusInter.Sa1010s on SC5.C5Cliente equals SA1.A1Cod
                                  join SA3 in ProtheusInter.Sa3010s on SC5.C5Vend1 equals SA3.A3Cod
                                  join SF4 in ProtheusInter.Sf4010s on SC6.C6Tes equals SF4.F4Codigo
                                  join SB1 in ProtheusInter.Sb1010s on SC6.C6Produto equals SB1.B1Cod
                                  join SUA in ProtheusInter.Sua010s on SC5.C5Uproces equals SUA.UaNum into Sr
                                  from c in Sr.DefaultIfEmpty()
                                  join SX5 in ProtheusInter.Sx5010s on new { Grupo = SA1.A1Xgrinte, Tabela = "Z3" } equals new { Grupo = SX5.X5Chave, Tabela = SX5.X5Tabela } into Se
                                  from a in Se.DefaultIfEmpty()
                                  where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num && a.DELET != "*"
                                  && SC6.C6Nota == ""
                                  && SC5.C5Nota == ""
                                  && SC6.C6Blq != "R"
                                  && SC6.DELET != "*"
                                  && SF4.F4Codigo == SC6.C6Tes
                                  && SF4.F4Duplic == "S"
                                  && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                                  && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                                  && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T")
                                  && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                                  && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                                  orderby SC5.C5Num, SC5.C5Emissao descending
                                  select new RelatorioCirurgiaAFaturar
                                  {
                                      Vendedor = SC5.C5Nomvend,
                                      Cliente = SC5.C5Nomcli,
                                      GrupoCliente = a.X5Descri,
                                      ClienteEntrega = SC5.C5Nomclie,
                                      Medico = SC5.C5XNmmed,
                                      Convenio = SC5.C5XNmpla,
                                      Cirurgia = Convert.ToInt32(SC5.C5XDtcir),
                                      Hoje = "",
                                      Dias = 0,
                                      Anging = "",
                                      Paciente = SC5.C5XNmpac,
                                      Matric = "",
                                      INPART = "",
                                      Valor = SC6.C6Valor,
                                      PVFaturamento = SC5.C5Num,
                                      Status = c.UaXstatus,
                                      DataEnvRA = " / / ",
                                      DataRecRA = " / / ",
                                      DataValorizacao = SC5.C5Xdtval,
                                      Emissao = Convert.ToInt32(SC5.C5Emissao)
                                  }
                        ).GroupBy(x => new
                        {
                            x.Vendedor,
                            x.Cliente,
                            x.GrupoCliente,
                            x.ClienteEntrega,
                            x.Medico,
                            x.Convenio,
                            x.Cirurgia,
                            x.Hoje,
                            x.Dias,
                            x.Anging,
                            x.Paciente,
                            x.Matric,
                            x.INPART,
                            x.PVFaturamento,
                            x.Status,
                            x.DataEnvRA,
                            x.DataRecRA,
                            x.DataValorizacao,
                            x.Emissao
                        })
                        .Select(x => new RelatorioCirurgiaAFaturar
                        {
                            Vendedor = x.Key.Vendedor,
                            Cliente = x.Key.Cliente,
                            GrupoCliente = x.Key.GrupoCliente,
                            ClienteEntrega = x.Key.ClienteEntrega,
                            Medico = x.Key.Medico,
                            Convenio = x.Key.Convenio,
                            Cirurgia = x.Key.Cirurgia,
                            Hoje = x.Key.Hoje,
                            Dias = x.Key.Dias,
                            Anging = x.Key.Anging,
                            Paciente = x.Key.Paciente,
                            Matric = x.Key.Matric,
                            INPART = x.Key.INPART,
                            Valor = x.Sum(c => c.Valor),
                            PVFaturamento = x.Key.PVFaturamento,
                            Status = x.Key.Status,
                            DataEnvRA = x.Key.DataEnvRA,
                            DataRecRA = x.Key.DataRecRA,
                            DataValorizacao = x.Key.DataValorizacao,
                            Empresa = "INTERMEDIC",
                            Emissao = x.Key.Emissao
                        }).ToList();
                #endregion

                Relatorio = Relatorio.Concat(RelatorioInter).ToList();

                Relatorio = Relatorio.Where(x => x.Status != "C").ToList();

                if (id == "1")
                {
                    DateTime ago = DateTime.Now.AddMonths(-3);

                    var Data = Convert.ToInt32(ago.ToString("yyyy/MM/dd").Replace("/", ""));

                    Relatorio = Relatorio.Where(x => x.Cirurgia >= Data).ToList();
                }
                else if (id == "2")
                {
                    DateTime ago = DateTime.Now.AddMonths(-9);
                    DateTime After = ago.AddMonths(6);

                    var Data = Convert.ToInt32(ago.ToString("yyyy/MM/dd").Replace("/", ""));
                    var Data2 = Convert.ToInt32(After.ToString("yyyy/MM/dd").Replace("/", ""));

                    Relatorio = Relatorio.Where(x => x.Cirurgia >= Data && x.Cirurgia < Data2).ToList();
                }
                else
                {
                    DateTime ago = DateTime.Now.AddMonths(-9);

                    var Data = Convert.ToInt32(ago.ToString("yyyy/MM/dd").Replace("/", ""));

                    Relatorio = Relatorio.Where(x => x.Cirurgia < Data).ToList();
                }

                Grupos = Relatorio.Where(x => x.GrupoCliente != "" && x.GrupoCliente != null).Select(x => x.GrupoCliente).Distinct().ToList();
                
                if (Grupo != null)
                {
                    Clientes = Relatorio.Where(x=> x.GrupoCliente == Grupo).Select(x => x.Cliente).Distinct().ToList();
                    this.Grupo = Grupo;
                    Relatorio = Relatorio.Where(x => x.GrupoCliente == Grupo).ToList();
                }

                Clientes = Relatorio.Select(x => x.Cliente).Distinct().ToList();
                
                if (NReduz != null)
                {
                    var registro = Relatorio.FirstOrDefault(x => x.Cliente == NReduz);
                    if (Grupo != null)
                    {
                        if (registro?.GrupoCliente == Grupo)
                        {
                            Cliente = NReduz;
                            Relatorio = Relatorio.Where(x => x.Cliente == NReduz).ToList();

                        }
                    }
                    else
                    {
                        Cliente = NReduz;
                        Relatorio = Relatorio.Where(x => x.Cliente == NReduz).ToList();
                    }
                }

                var data = DateTime.Now;
                this.Empresa = Empresa;
                if (Empresa != null)
                {
                    Relatorio = Relatorio.Where(x => x.Empresa == Empresa).ToList();
                }

                Relatorio.ForEach(x =>
                {
                    x.Hoje = data.ToString("dd/MM/yyyy");

                    var check = x.Cirurgia;

                    if (check != 0)
                    {
                        x.Dias = (int)(data - Convert.ToDateTime($"{x.Cirurgia.ToString().Substring(4, 2)}/{x.Cirurgia.ToString().Substring(6, 2)}/{x.Cirurgia.ToString().Substring(0, 4)}")).TotalDays;
                    }

                    var check2 = int.TryParse(x.DataValorizacao, out int result2);

                    if (check2)
                    {
                        x.DataValorizacao = $"{x.DataValorizacao.Substring(6, 2)}/{x.DataValorizacao.Substring(4, 2)}/{x.DataValorizacao.Substring(0, 4)}";
                    }


                    if (x.Dias <= 30)
                    {
                        x.Anging = "ATE 30";
                    }
                    else if (x.Dias <= 60)
                    {
                        x.Anging = "31 A 60";
                    }
                    else if (x.Dias <= 90)
                    {
                        x.Anging = "61 A 90";
                    }
                    else if (x.Dias <= 120)
                    {
                        x.Anging = "91 A 120";
                    }
                    else
                    {
                        x.Anging = "MAIS DE 120";
                    }
                    Total += x.Valor;
                });

                Relatorio = Relatorio.OrderBy(x => x.Dias).ToList();

            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiaPorTempo", user);
            }

            return Page();
        }

        public IActionResult OnPostExport(string id, string NReduz, string Empresa, string Grupo)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();

                #region Relatorio
                Relatorio = (from SC5 in Protheus.Sc5010s
                             join SC6 in Protheus.Sc6010s on new { Filial = SC5.C5Filial, Num = SC5.C5Num } equals new { Filial = SC6.C6Filial, Num = SC6.C6Num }
                             join SA1 in Protheus.Sa1010s on SC5.C5Cliente equals SA1.A1Cod
                             join SA3 in Protheus.Sa3010s on SC5.C5Vend1 equals SA3.A3Cod
                             join SF4 in Protheus.Sf4010s on SC6.C6Tes equals SF4.F4Codigo
                             join SB1 in Protheus.Sb1010s on SC6.C6Produto equals SB1.B1Cod
                             join SUA in Protheus.Sua010s on SC5.C5Uproces equals SUA.UaNum into Sr
                             from c in Sr.DefaultIfEmpty()
                             join SX5 in Protheus.Sx5010s on new { Grupo = SA1.A1Xgrinte, Filial = SC5.C5Filial, Tabela = "Z3" } equals new { Grupo = SX5.X5Chave, Filial = SX5.X5Filial, Tabela = SX5.X5Tabela } into Se
                             from a in Se.DefaultIfEmpty()
                             where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num && a.DELET != "*"
                             && SC6.C6Nota == ""
                             && SC5.C5Nota == ""
                             && SC6.C6Blq != "R"
                             && SC6.DELET != "*"
                             && SF4.F4Codigo == SC6.C6Tes
                             && SF4.F4Duplic == "S"
                             && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                             && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                             && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T")
                             && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                             && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                             orderby SC5.C5Num, SC5.C5Emissao descending
                             select new RelatorioCirurgiaAFaturar
                             {
                                 Vendedor = SC5.C5Nomvend,
                                 Cliente = SC5.C5Nomcli,
                                 GrupoCliente = a.X5Descri,
                                 ClienteEntrega = SC5.C5Nomclie,
                                 Medico = SC5.C5XNmmed,
                                 Convenio = SC5.C5XNmpla,
                                 Cirurgia = Convert.ToInt32(SC5.C5XDtcir),
                                 Hoje = "",
                                 Dias = 0,
                                 Anging = "",
                                 Paciente = SC5.C5XNmpac,
                                 Matric = "",
                                 INPART = "",
                                 Valor = SC6.C6Valor,
                                 PVFaturamento = SC5.C5Num,
                                 Status = c.UaXstatus,
                                 DataEnvRA = " / / ",
                                 DataRecRA = " / / ",
                                 DataValorizacao = SC5.C5Xdtval,
                                 Emissao = Convert.ToInt32(SC5.C5Emissao)
                             }
                        ).GroupBy(x => new
                        {
                            x.Vendedor,
                            x.Cliente,
                            x.GrupoCliente,
                            x.ClienteEntrega,
                            x.Medico,
                            x.Convenio,
                            x.Cirurgia,
                            x.Hoje,
                            x.Dias,
                            x.Anging,
                            x.Paciente,
                            x.Matric,
                            x.INPART,
                            x.PVFaturamento,
                            x.Status,
                            x.DataEnvRA,
                            x.DataRecRA,
                            x.DataValorizacao,
                            x.Emissao
                        })
                        .Select(x => new RelatorioCirurgiaAFaturar
                        {
                            Vendedor = x.Key.Vendedor,
                            Cliente = x.Key.Cliente,
                            GrupoCliente = x.Key.GrupoCliente,
                            ClienteEntrega = x.Key.ClienteEntrega,
                            Medico = x.Key.Medico,
                            Convenio = x.Key.Convenio,
                            Cirurgia = x.Key.Cirurgia,
                            Hoje = x.Key.Hoje,
                            Dias = x.Key.Dias,
                            Anging = x.Key.Anging,
                            Paciente = x.Key.Paciente,
                            Matric = x.Key.Matric,
                            INPART = x.Key.INPART,
                            Valor = x.Sum(c => c.Valor),
                            PVFaturamento = x.Key.PVFaturamento,
                            Status = x.Key.Status,
                            DataEnvRA = x.Key.DataEnvRA,
                            DataRecRA = x.Key.DataRecRA,
                            DataValorizacao = x.Key.DataValorizacao,
                            Empresa = "DENUO",
                            Emissao = x.Key.Emissao
                        }).ToList();

                RelatorioInter = (from SC5 in ProtheusInter.Sc5010s
                                  join SC6 in ProtheusInter.Sc6010s on new { Filial = SC5.C5Filial, Num = SC5.C5Num } equals new { Filial = SC6.C6Filial, Num = SC6.C6Num }
                                  join SA1 in ProtheusInter.Sa1010s on SC5.C5Cliente equals SA1.A1Cod
                                  join SA3 in ProtheusInter.Sa3010s on SC5.C5Vend1 equals SA3.A3Cod
                                  join SF4 in ProtheusInter.Sf4010s on SC6.C6Tes equals SF4.F4Codigo
                                  join SB1 in ProtheusInter.Sb1010s on SC6.C6Produto equals SB1.B1Cod
                                  join SUA in ProtheusInter.Sua010s on SC5.C5Uproces equals SUA.UaNum into Sr
                                  from c in Sr.DefaultIfEmpty()
                                  join SX5 in ProtheusInter.Sx5010s on new { Grupo = SA1.A1Xgrinte, Tabela = "Z3" } equals new { Grupo = SX5.X5Chave, Tabela = SX5.X5Tabela } into Se
                                  from a in Se.DefaultIfEmpty()
                                  where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num && a.DELET != "*"
                                  && SC6.C6Nota == ""
                                  && SC5.C5Nota == ""
                                  && SC6.C6Blq != "R"
                                  && SC6.DELET != "*"
                                  && SF4.F4Codigo == SC6.C6Tes
                                  && SF4.F4Duplic == "S"
                                  && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                                  && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                                  && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T")
                                  && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                                  && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140"
                                  orderby SC5.C5Num, SC5.C5Emissao descending
                                  select new RelatorioCirurgiaAFaturar
                                  {
                                      Vendedor = SC5.C5Nomvend,
                                      Cliente = SC5.C5Nomcli,
                                      GrupoCliente = a.X5Descri,
                                      ClienteEntrega = SC5.C5Nomclie,
                                      Medico = SC5.C5XNmmed,
                                      Convenio = SC5.C5XNmpla,
                                      Cirurgia = Convert.ToInt32(SC5.C5XDtcir),
                                      Hoje = "",
                                      Dias = 0,
                                      Anging = "",
                                      Paciente = SC5.C5XNmpac,
                                      Matric = "",
                                      INPART = "",
                                      Valor = SC6.C6Valor,
                                      PVFaturamento = SC5.C5Num,
                                      Status = c.UaXstatus,
                                      DataEnvRA = " / / ",
                                      DataRecRA = " / / ",
                                      DataValorizacao = SC5.C5Xdtval,
                                      Emissao = Convert.ToInt32(SC5.C5Emissao)
                                  }
                        ).GroupBy(x => new
                        {
                            x.Vendedor,
                            x.Cliente,
                            x.GrupoCliente,
                            x.ClienteEntrega,
                            x.Medico,
                            x.Convenio,
                            x.Cirurgia,
                            x.Hoje,
                            x.Dias,
                            x.Anging,
                            x.Paciente,
                            x.Matric,
                            x.INPART,
                            x.PVFaturamento,
                            x.Status,
                            x.DataEnvRA,
                            x.DataRecRA,
                            x.DataValorizacao,
                            x.Emissao
                        })
                        .Select(x => new RelatorioCirurgiaAFaturar
                        {
                            Vendedor = x.Key.Vendedor,
                            Cliente = x.Key.Cliente,
                            GrupoCliente = x.Key.GrupoCliente,
                            ClienteEntrega = x.Key.ClienteEntrega,
                            Medico = x.Key.Medico,
                            Convenio = x.Key.Convenio,
                            Cirurgia = x.Key.Cirurgia,
                            Hoje = x.Key.Hoje,
                            Dias = x.Key.Dias,
                            Anging = x.Key.Anging,
                            Paciente = x.Key.Paciente,
                            Matric = x.Key.Matric,
                            INPART = x.Key.INPART,
                            Valor = x.Sum(c => c.Valor),
                            PVFaturamento = x.Key.PVFaturamento,
                            Status = x.Key.Status,
                            DataEnvRA = x.Key.DataEnvRA,
                            DataRecRA = x.Key.DataRecRA,
                            DataValorizacao = x.Key.DataValorizacao,
                            Empresa = "INTERMEDIC",
                            Emissao = x.Key.Emissao
                        }).ToList();
                #endregion

                Relatorio = Relatorio.Concat(RelatorioInter).ToList();

                Relatorio = Relatorio.Where(x => x.Status != "C").ToList();

                if (id == "1")
                {
                    DateTime ago = DateTime.Now.AddMonths(-3);

                    var Data = Convert.ToInt32(ago.ToString("yyyy/MM/dd").Replace("/", ""));

                    Relatorio = Relatorio.Where(x => x.Cirurgia >= Data).ToList();
                }
                else if (id == "2")
                {
                    DateTime ago = DateTime.Now.AddMonths(-9);
                    DateTime After = ago.AddMonths(6);

                    var Data = Convert.ToInt32(ago.ToString("yyyy/MM/dd").Replace("/", ""));
                    var Data2 = Convert.ToInt32(After.ToString("yyyy/MM/dd").Replace("/", ""));

                    Relatorio = Relatorio.Where(x => x.Cirurgia >= Data && x.Cirurgia < Data2).ToList();
                }
                else
                {
                    DateTime ago = DateTime.Now.AddMonths(-9);

                    var Data = Convert.ToInt32(ago.ToString("yyyy/MM/dd").Replace("/", ""));

                    Relatorio = Relatorio.Where(x => x.Cirurgia < Data).ToList();
                }

                this.Grupo = Grupo;
                if (Grupo != null)
                {
                    Relatorio = Relatorio.Where(x => x.GrupoCliente == Grupo).ToList();
                }

                Cliente = NReduz;
                if (NReduz != null)
                {
                    var registro = Relatorio.FirstOrDefault(x => x.Cliente == NReduz);
                    if (Grupo != null)
                    {
                        if (registro?.GrupoCliente == Grupo)
                        {
                            Cliente = NReduz;
                            Relatorio = Relatorio.Where(x => x.Cliente == NReduz).ToList();

                        }
                    }
                    else
                    {
                        Cliente = NReduz;
                        Relatorio = Relatorio.Where(x => x.Cliente == NReduz).ToList();
                    }
                }

                

                var data = DateTime.Now;
                this.Empresa = Empresa;
                if (Empresa != null)
                {
                    Relatorio = Relatorio.Where(x => x.Empresa == Empresa).ToList();
                }

                Relatorio.ForEach(x =>
                {
                    x.Hoje = data.ToString("dd/MM/yyyy");

                    var check = x.Cirurgia;

                    if (check != 0)
                    {
                        x.Dias = (int)(data - Convert.ToDateTime($"{x.Cirurgia.ToString().Substring(4, 2)}/{x.Cirurgia.ToString().Substring(6, 2)}/{x.Cirurgia.ToString().Substring(0, 4)}")).TotalDays;
                    }
                    

                    var check2 = int.TryParse(x.DataValorizacao, out int result2);

                    if (check2)
                    {
                        x.DataValorizacao = $"{x.DataValorizacao.Substring(6, 2)}/{x.DataValorizacao.Substring(4, 2)}/{x.DataValorizacao.Substring(0, 4)}";
                    }


                    if (x.Dias <= 30)
                    {
                        x.Anging = "ATE 30";
                    }
                    else if (x.Dias <= 60)
                    {
                        x.Anging = "31 A 60";
                    }
                    else if (x.Dias <= 90)
                    {
                        x.Anging = "61 A 90";
                    }
                    else if (x.Dias <= 120)
                    {
                        x.Anging = "91 A 120";
                    }
                    else
                    {
                        x.Anging = "MAIS DE 120";
                    }
                    Total += x.Valor;
                });

                Relatorio = Relatorio.OrderBy(x => x.Dias).ToList();
                

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Cirurgias A Faturar");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Cirurgias A Faturar");

                sheet.Cells[1, 1].Value = "VENDEDOR";
                sheet.Cells[1, 2].Value = "CLIENTE FAT";
                sheet.Cells[1, 3].Value = "GRUPO CLIENTE";
                sheet.Cells[1, 4].Value = "CLIENTE ENTR";
                sheet.Cells[1, 5].Value = "MEDICO";
                sheet.Cells[1, 6].Value = "CONVENIO";
                sheet.Cells[1, 7].Value = "CIRURGIA";
                sheet.Cells[1, 8].Value = "HOJE";
                sheet.Cells[1, 9].Value = "DIAS";
                sheet.Cells[1, 10].Value = "ANGING";
                sheet.Cells[1, 11].Value = "PACIENTE";
                sheet.Cells[1, 12].Value = "MATRIC";
                sheet.Cells[1, 13].Value = "INPART";
                sheet.Cells[1, 14].Value = "VALOR TOTAL";
                sheet.Cells[1, 15].Value = "PVFATURA";
                sheet.Cells[1, 16].Value = "STATUS";
                sheet.Cells[1, 17].Value = "DATA ENV. RAH";
                sheet.Cells[1, 18].Value = "DATA REC. RAH";
                sheet.Cells[1, 19].Value = "DATA VALORIZAÇÃO";
                sheet.Cells[1, 20].Value = "EMPRESA";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Vendedor;
                    sheet.Cells[i, 2].Value = Pedido.Cliente;
                    sheet.Cells[i, 3].Value = Pedido.GrupoCliente;
                    sheet.Cells[i, 4].Value = Pedido.ClienteEntrega;
                    sheet.Cells[i, 5].Value = Pedido.Medico;
                    sheet.Cells[i, 6].Value = Pedido.Convenio;
                    sheet.Cells[i, 7].Value = Pedido.Cirurgia;
                    sheet.Cells[i, 8].Value = Pedido.Hoje;
                    sheet.Cells[i, 9].Value = Pedido.Dias;
                    sheet.Cells[i, 10].Value = Pedido.Anging;
                    sheet.Cells[i, 11].Value = Pedido.Paciente;
                    sheet.Cells[i, 12].Value = Pedido.Matric;
                    sheet.Cells[i, 13].Value = Pedido.INPART;
                    sheet.Cells[i, 14].Value = Pedido.Valor;
                    sheet.Cells[i, 15].Value = Pedido.PVFaturamento;
                    sheet.Cells[i, 16].Value = Pedido.Status;
                    sheet.Cells[i, 17].Value = Pedido.DataEnvRA;
                    sheet.Cells[i, 18].Value = Pedido.DataRecRA;
                    sheet.Cells[i, 19].Value = Pedido.DataValorizacao;
                    sheet.Cells[i, 20].Value = Pedido.Empresa;
                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CirurgiaAFaturar.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiaAFaturarADM Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
