using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.AdmVendas;

namespace SGID.Pages.Relatorios.RH
{

    [Authorize]
    public class CirurgiasValorizadasInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioCirurgiasValorizadas> Relatorio = new List<RelatorioCirurgiasValorizadas>();
        public CirurgiasValorizadasInterModel(TOTVSINTERContext context,ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public double QtdPedido { get; set; }
        public double QtdFaturada { get; set; }
        public double VlUnitario { get; set; }
        public double TotalPedido { get; set; }
        public double TotalFat { get; set; }
        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(DateTime DataInicio,DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                Relatorio = (from SC50 in Protheus.Sc5010s
                             join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                             join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             join PA10 in Protheus.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                             from c in Sr.DefaultIfEmpty()
                             join SUA10 in Protheus.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                             from a in Rs.DefaultIfEmpty()
                             where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                             && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                             orderby SC50.C5XDtcir descending
                             select new RelatorioCirurgiasValorizadas
                             {
                                 DTCirurgia = $"{SC50.C5XDtcir.Substring(6, 2)}/{SC50.C5XDtcir.Substring(4, 2)}/{SC50.C5XDtcir.Substring(0, 4)}",
                                 Filial = SC50.C5Filial,
                                 NumPedido = SC50.C5Num,
                                 Agendamento = a.UaUnumage,
                                 TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC50.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                                 Vendedor = SC50.C5Nomvend,
                                 Medico = SC50.C5XNmmed,
                                 Paciente = SC50.C5XNmpac,
                                 ClienteEnt = SC50.C5XNmpla,
                                 NumPatrim = SC60.C6Upatrim,
                                 Patrimonio = c.Pa1Despat ?? "SEM PATRIMONIO",
                                 Produto = SC60.C6Produto,
                                 DescProd = SB10.B1Desc,
                                 QTDPedido = SC60.C6Qtdven,
                                 QTDFaturada = SC60.C6Qtdent,
                                 VLUnitario = SC60.C6Prcven,
                                 TotalPedido = SC60.C6Valor,
                                 TotalFat = SC60.C6Qtdent * SC60.C6Prcven,
                                 Faturado = SC50.C5Nota != "" || (SC50.C5Liberok == "E" && SC50.C5Blq == "") ? "S" : "N"
                             }
                             ).OrderBy(x => x.Vendedor).ToList();

                Relatorio.ForEach(Pedido =>
                {
                    QtdPedido += Pedido.QTDPedido;
                    QtdFaturada += Pedido.QTDFaturada;
                    VlUnitario += Pedido.VLUnitario;
                    TotalPedido += Pedido.TotalPedido;
                    TotalFat += Pedido.TotalFat;
                });

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasValorizadasInterRH",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Relatorio = (from SC50 in Protheus.Sc5010s
                             join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                             join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             join PA10 in Protheus.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                             from c in Sr.DefaultIfEmpty()
                             join SUA10 in Protheus.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                             from a in Rs.DefaultIfEmpty()
                             where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                             && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                             orderby SC50.C5XDtcir descending
                             select new RelatorioCirurgiasValorizadas
                             {
                                 DTCirurgia = $"{SC50.C5XDtcir.Substring(6, 2)}/{SC50.C5XDtcir.Substring(4, 2)}/{SC50.C5XDtcir.Substring(0, 4)}",
                                 Filial = SC50.C5Filial,
                                 NumPedido = SC50.C5Num,
                                 Agendamento = a.UaUnumage,
                                 TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC50.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                                 Vendedor = SC50.C5Nomvend,
                                 Medico = SC50.C5XNmmed,
                                 Paciente = SC50.C5XNmpac,
                                 ClienteEnt = SC50.C5XNmpla,
                                 NumPatrim = SC60.C6Upatrim,
                                 Patrimonio = c.Pa1Despat ?? "SEM PATRIMONIO",
                                 Produto = SC60.C6Produto,
                                 DescProd = SB10.B1Desc,
                                 QTDPedido = SC60.C6Qtdven,
                                 QTDFaturada = SC60.C6Qtdent,
                                 VLUnitario = SC60.C6Prcven,
                                 TotalPedido = SC60.C6Valor,
                                 TotalFat = SC60.C6Qtdent * SC60.C6Prcven,
                                 Faturado = SC50.C5Nota != "" || (SC50.C5Liberok == "E" && SC50.C5Blq == "") ? "S" : "N"
                             }
                             ).OrderBy(x => x.Vendedor).ToList();

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
                sheet.Cells[1, 11].Value = "Num. Patrim.";
                sheet.Cells[1, 12].Value = "Patrimonio";
                sheet.Cells[1, 13].Value = "Produto";
                sheet.Cells[1, 14].Value = "Desc. Produto";
                sheet.Cells[1, 15].Value = "QTD Pedido";
                sheet.Cells[1, 16].Value = "QTD Faturada";
                sheet.Cells[1, 17].Value = "VL Unitário";
                sheet.Cells[1, 18].Value = "Total Pedido";
                sheet.Cells[1, 19].Value = "Total Fat.";
                sheet.Cells[1, 20].Value = "Faturado ?";

                int i = 2;

                QtdPedido = 0;
                QtdFaturada = 0;
                VlUnitario = 0;
                TotalPedido = 0;
                TotalFat = 0;

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
                    sheet.Cells[i, 13].Value = Pedido.Produto;
                    sheet.Cells[i, 14].Value = Pedido.DescProd;
                    sheet.Cells[i, 15].Value = Pedido.QTDPedido;
                    sheet.Cells[i, 16].Value = Pedido.QTDFaturada;
                    sheet.Cells[i, 17].Value = Pedido.VLUnitario;
                    sheet.Cells[i, 18].Value = Pedido.TotalPedido;
                    sheet.Cells[i, 19].Value = Pedido.TotalFat;
                    sheet.Cells[i, 20].Value = Pedido.Faturado;

                    QtdPedido += Pedido.QTDPedido;
                    QtdFaturada += Pedido.QTDFaturada;
                    VlUnitario += Pedido.VLUnitario;
                    TotalPedido += Pedido.TotalPedido;
                    TotalFat += Pedido.TotalFat;

                    i++;
                });

                sheet.Cells[i, 1].Value = "Total Geral";

                sheet.Cells[i, 15].Value = QtdPedido;
                sheet.Cells[i, 16].Value = QtdFaturada;
                sheet.Cells[i, 17].Value = VlUnitario;
                sheet.Cells[i, 18].Value = TotalPedido;
                sheet.Cells[i, 19].Value = TotalFat;

                sheet.Cells[i, 1, i, 14].Merge = true;

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CirurgiasValorizadasInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasValorizadasInterRH Excel",user);

                return LocalRedirect("/error");
            }
        }
    }
}
