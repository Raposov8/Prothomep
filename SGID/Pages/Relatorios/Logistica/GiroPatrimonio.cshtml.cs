using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.AdmVendas;
using SGID.Models.Denuo;

namespace SGID.Pages.Relatorios.Logistica
{
    public class GiroPatrimonioModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioCirurgiasValorizadas> Relatorio = new List<RelatorioCirurgiasValorizadas>();

        public GiroPatrimonioModel(TOTVSDENUOContext protheus, ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public DateTime Inicio { get; set; } = DateTime.Now;
        public DateTime Fim { get; set; } = DateTime.Now;
        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                string user = User.Identity.Name.Split("@")[0].ToUpper();

                var relatorio = (from SC50 in Protheus.Sc5010s
                                 join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                 join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                 join SB10 in Protheus.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                 join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                 join PA10 in Protheus.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                 from c in Sr.DefaultIfEmpty()
                                 join SUA10 in Protheus.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                 from a in Rs.DefaultIfEmpty()
                                 where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                 && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                                 && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                                 && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                 select new RelatorioCirurgiasValorizadas
                                 {
                                     DTCirurgia = SC50.C5XDtcir,
                                     Filial = SC50.C5Filial,
                                     NumPedido = SC50.C5Num,
                                     Agendamento = a.UaUnumage,
                                     TipoOper = SC50.C5Utpoper,
                                     Vendedor = SC50.C5Nomvend,
                                     Medico = SC50.C5XNmmed,
                                     Paciente = SC50.C5XNmpac,
                                     ClienteEnt = SC50.C5XNmpla,
                                     NumPatrim = SC60.C6Upatrim,
                                     Patrimonio = c.Pa1Despat ?? "SEM PATRIMONIO"
                                 }
                                 ).
                                 GroupBy(x => new
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
                                     x.NumPatrim,
                                     x.Patrimonio
                                 });


                Relatorio = relatorio.Select(x => new RelatorioCirurgiasValorizadas
                {
                    DTCirurgia = $"{x.Key.DTCirurgia.Substring(6, 2)}/{x.Key.DTCirurgia.Substring(4, 2)}/{x.Key.DTCirurgia.Substring(0, 4)}",
                    Filial = x.Key.Filial,
                    NumPedido = x.Key.NumPedido,
                    Agendamento = x.Key.Agendamento,
                    TipoOper = x.Key.TipoOper == "F" ? "FATURAMENTO DE CIRURGIAS" : x.Key.TipoOper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                    Vendedor = x.Key.Vendedor,
                    Medico = x.Key.Medico,
                    Paciente = x.Key.Paciente,
                    ClienteEnt = x.Key.ClienteEnt,
                    NumPatrim = x.Key.NumPatrim,
                    Patrimonio = x.Key.Patrimonio
                }).ToList();


                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "GiroPatrimonioInter", user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {

                Inicio = DataInicio;
                Fim = DataFim;

                string user = User.Identity.Name.Split("@")[0].ToUpper();

                var relatorio = (from SC50 in Protheus.Sc5010s
                                 join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                 join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                 join SB10 in Protheus.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                 join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                 join PA10 in Protheus.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                                 from c in Sr.DefaultIfEmpty()
                                 join SUA10 in Protheus.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                                 from a in Rs.DefaultIfEmpty()
                                 where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                                 && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                                 && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                                 && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                 select new RelatorioCirurgiasValorizadas
                                 {
                                     DTCirurgia = SC50.C5XDtcir,
                                     Filial = SC50.C5Filial,
                                     NumPedido = SC50.C5Num,
                                     Agendamento = a.UaUnumage,
                                     TipoOper = SC50.C5Utpoper,
                                     Vendedor = SC50.C5Nomvend,
                                     Medico = SC50.C5XNmmed,
                                     Paciente = SC50.C5XNmpac,
                                     ClienteEnt = SC50.C5XNmpla,
                                     NumPatrim = SC60.C6Upatrim,
                                     Patrimonio = c.Pa1Despat ?? "SEM PATRIMONIO"
                                 }
                              ).
                              GroupBy(x => new
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
                                  x.NumPatrim,
                                  x.Patrimonio
                              });


                Relatorio = relatorio.Select(x => new RelatorioCirurgiasValorizadas
                {
                    DTCirurgia = $"{x.Key.DTCirurgia.Substring(6, 2)}/{x.Key.DTCirurgia.Substring(4, 2)}/{x.Key.DTCirurgia.Substring(0, 4)}",
                    Filial = x.Key.Filial,
                    NumPedido = x.Key.NumPedido,
                    Agendamento = x.Key.Agendamento,
                    TipoOper = x.Key.TipoOper == "F" ? "FATURAMENTO DE CIRURGIAS" : x.Key.TipoOper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                    Vendedor = x.Key.Vendedor,
                    Medico = x.Key.Medico,
                    Paciente = x.Key.Paciente,
                    ClienteEnt = x.Key.ClienteEnt,
                    NumPatrim = x.Key.NumPatrim,
                    Patrimonio = x.Key.Patrimonio
                }).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Giro Patrimonio Inter");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Giro Patrimonio Inter");

                sheet.Cells[1, 1].Value = "DATA CIRURGIA";
                sheet.Cells[1, 2].Value = "FILIAL";
                sheet.Cells[1, 3].Value = "NUM. PEDIDO";
                sheet.Cells[1, 4].Value = "AGENDAMENTO";
                sheet.Cells[1, 5].Value = "TIPO";
                sheet.Cells[1, 6].Value = "VENDEDOR";
                sheet.Cells[1, 7].Value = "MÉDICO";
                sheet.Cells[1, 8].Value = "PACIENTE";
                sheet.Cells[1, 9].Value = "CLIENTE ENTREGA";
                sheet.Cells[1, 10].Value = "CONVÊNIO";
                sheet.Cells[1, 11].Value = "NUM. PATRIM.";
                sheet.Cells[1, 12].Value = "PATRIMONIO";

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
                    sheet.Cells[i, 11].Value = Pedido.NumPatrim;
                    sheet.Cells[i, 12].Value = Pedido.Patrimonio;

                    i++;
                });


                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "GiroPatrimonioInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "GiroPatrimonioInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
