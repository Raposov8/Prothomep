using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.AdmVendas;
using SGID.Models.Denuo;
using SGID.Models.Estoque.RelatorioFaturamentoNFFab;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.AdmVendas
{
    public class CirurgiaAFaturarInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public List<RelatorioCirurgiaAFaturar> Relatorio { get; set; } = new List<RelatorioCirurgiaAFaturar>();
        public List<string> Pedidos { get; set; } = new List<string>();
        public List<string> Clientes { get; set; } = new List<string>();
        public List<RelatorioCirurgiaAFaturar> Draw1 { get; set; } = new List<RelatorioCirurgiaAFaturar>();
        public double Total { get; set; }
        public string Anging { get; set; }
        public string Cliente { get; set; }
        public CirurgiaAFaturarInterModel(TOTVSINTERContext context, ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }

        public void OnGet()
        {
            try
            {

                string user = User.Identity.Name.Split("@")[0].ToUpper();

                Relatorio = (from SC5 in Protheus.Sc5010s
                             from SC6 in Protheus.Sc6010s
                             from SA1 in Protheus.Sa1010s
                             from SA3 in Protheus.Sa3010s
                             from SF4 in Protheus.Sf4010s
                             from SB1 in Protheus.Sb1010s
                             where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                             && SC6.C6Nota == ""
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
                                 GrupoCliente = "",
                                 ClienteEntrega = SC5.C5Nomclie,
                                 Medico = SC5.C5XNmmed,
                                 Convenio = SC5.C5XNmpla,
                                 Cirurgia = SC5.C5XDtcir,
                                 Hoje = "",
                                 Dias = 0,
                                 Anging = "",
                                 Paciente = SC5.C5XNmpac,
                                 Matric = "",
                                 INPART = "",
                                 Valor = SC6.C6Valor,
                                 PVFaturamento = SC5.C5Num,
                                 Status = "",
                                 DataEnvRA = " / / ",
                                 DataRecRA = " / / ",
                                 DataValorizacao = SC5.C5Xdtval
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
                                x.DataValorizacao
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
                                DataValorizacao = x.Key.DataValorizacao

                            }).ToList();

                Clientes = Relatorio.Select(x => x.Cliente).Distinct().ToList();

                var data = DateTime.Now;

                Relatorio.ForEach(x =>
                {
                    x.Hoje = data.ToString("dd/MM/yyyy");

                    var check = int.TryParse(x.Cirurgia, out int result);

                    if (check)
                    {
                        x.Dias = (int)(data - Convert.ToDateTime($"{x.Cirurgia.Substring(4, 2)}/{x.Cirurgia.Substring(6, 2)}/{x.Cirurgia.Substring(0, 4)}")).TotalDays;

                        x.Cirurgia = $"{x.Cirurgia.Substring(6, 2)}/{x.Cirurgia.Substring(4, 2)}/{x.Cirurgia.Substring(0, 4)}";
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

                Draw1 = Relatorio.GroupBy(x => x.Cliente).Select(x => new RelatorioCirurgiaAFaturar { Cliente = x.Key, Valor = x.Sum(c => c.Valor) }).ToList();

                Relatorio = Relatorio.OrderBy(x => x.Dias).ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasAFaturarRH", user);
            }
        }

        public IActionResult OnPost(string Cliente, string Tempo)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();

                Relatorio = (from SC5 in Protheus.Sc5010s
                             from SC6 in Protheus.Sc6010s
                             from SA1 in Protheus.Sa1010s
                             from SA3 in Protheus.Sa3010s
                             from SF4 in Protheus.Sf4010s
                             from SB1 in Protheus.Sb1010s
                             where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                             && SC6.C6Nota == ""
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
                                 GrupoCliente = "",
                                 ClienteEntrega = SC5.C5Nomclie,
                                 Medico = SC5.C5XNmmed,
                                 Convenio = SC5.C5XNmpla,
                                 Cirurgia = SC5.C5XDtcir,
                                 Hoje = "",
                                 Dias = 0,
                                 Anging = "",
                                 Paciente = SC5.C5XNmpac,
                                 Matric = "",
                                 INPART = "",
                                 Valor = SC6.C6Valor,
                                 PVFaturamento = SC5.C5Num,
                                 Status = "",
                                 DataEnvRA = " / / ",
                                 DataRecRA = " / / ",
                                 DataValorizacao = SC5.C5Xdtval
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
                                x.DataValorizacao
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
                                DataValorizacao = x.Key.DataValorizacao

                            }).ToList();

                Clientes = Relatorio.Select(x => x.Cliente).Distinct().ToList();

                if (Cliente != null && Cliente != "")
                {
                    Relatorio = Relatorio.Where(x => x.Cliente == Cliente).ToList();
                    this.Cliente = Cliente;
                }

                var data = DateTime.Now;

                Relatorio.ForEach(x =>
                {
                    x.Hoje = data.ToString("dd/MM/yyyy");

                    var check = int.TryParse(x.Cirurgia, out int result);

                    if (check)
                    {
                        x.Dias = (int)(data - Convert.ToDateTime($"{x.Cirurgia.Substring(4, 2)}/{x.Cirurgia.Substring(6, 2)}/{x.Cirurgia.Substring(0, 4)}")).TotalDays;

                        x.Cirurgia = $"{x.Cirurgia.Substring(6, 2)}/{x.Cirurgia.Substring(4, 2)}/{x.Cirurgia.Substring(0, 4)}";
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

                if (Tempo != null && Tempo != "")
                {
                    Relatorio = Relatorio.Where(x => x.Anging == Tempo).ToList();
                    Anging = Tempo;
                }

                Draw1 = Relatorio.GroupBy(x => x.Cliente).Select(x => new RelatorioCirurgiaAFaturar { Cliente = x.Key, Valor = x.Sum(c => c.Valor) }).ToList();

                Relatorio = Relatorio.OrderBy(x => x.Dias).ToList();
                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasAFaturarRH", user);
                return Page();
            }
        }

        public IActionResult OnPostExport(string Cliente, string Tempo)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();

                Relatorio = (from SC5 in Protheus.Sc5010s
                             from SC6 in Protheus.Sc6010s
                             from SA1 in Protheus.Sa1010s
                             from SA3 in Protheus.Sa3010s
                             from SF4 in Protheus.Sf4010s
                             from SB1 in Protheus.Sb1010s
                             where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                             && SC6.C6Nota == ""
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
                                 GrupoCliente = "",
                                 ClienteEntrega = SC5.C5Nomclie,
                                 Medico = SC5.C5XNmmed,
                                 Convenio = SC5.C5XNmpla,
                                 Cirurgia = SC5.C5XDtcir,
                                 Hoje = "",
                                 Dias = 0,
                                 Anging = "",
                                 Paciente = SC5.C5XNmpac,
                                 Matric = "",
                                 INPART = "",
                                 Valor = SC6.C6Valor,
                                 PVFaturamento = SC5.C5Num,
                                 Status = "",
                                 DataEnvRA = " / / ",
                                 DataRecRA = " / / ",
                                 DataValorizacao = SC5.C5Xdtval
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
                                x.DataValorizacao
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
                                DataValorizacao = x.Key.DataValorizacao

                            }).ToList();

                Clientes = Relatorio.Select(x => x.Cliente).Distinct().ToList();

                if (Cliente != null && Cliente != "")
                {
                    Relatorio = Relatorio.Where(x => x.Cliente == Cliente).ToList();
                }

                var data = DateTime.Now;

                Relatorio.ForEach(x =>
                {
                    x.Hoje = data.ToString("dd/MM/yyyy");

                    var check = int.TryParse(x.Cirurgia, out int result);

                    if (check)
                    {
                        x.Dias = (int)(data - Convert.ToDateTime($"{x.Cirurgia.Substring(4, 2)}/{x.Cirurgia.Substring(6, 2)}/{x.Cirurgia.Substring(0, 4)}")).TotalDays;

                        x.Cirurgia = $"{x.Cirurgia.Substring(6, 2)}/{x.Cirurgia.Substring(4, 2)}/{x.Cirurgia.Substring(0, 4)}";
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

                if (Tempo != null && Tempo != "")
                {
                    Relatorio = Relatorio.Where(x => x.Anging == Tempo).ToList();
                }

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
                    i++;
                });



                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CirurgiaAFaturarInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiaAFaturarADMInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
