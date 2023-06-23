using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Denuo;

namespace SGID.Pages.Relatorios.Cirurgias
{
    [Authorize]
    public class CirurgiasValorizadasModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioCirurgiasValorizadas> Relatorio = new List<RelatorioCirurgiasValorizadas>();

        public CirurgiasValorizadasModel(TOTVSDENUOContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }

        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }
        public double Pedido { get; set; }
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
                Pedido = 0;
                QtdFaturada = 0;
                VlUnitario = 0;
                TotalPedido = 0;
                TotalFat = 0;

                string user = User.Identity.Name.Split("@")[0].ToUpper();

                Relatorio = (from SC50 in Protheus.Sc5010s
                             join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                             join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             join PA10 in Protheus.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                             from c in Sr.DefaultIfEmpty()
                             join SUA10 in Protheus.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                             from a in Rs.DefaultIfEmpty()
                             where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                             && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             && (SA30.A3Xlogin == user || SA30.A3Xlogsup == user) && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                             orderby SC50.C5XDtcir descending
                             select new RelatorioCirurgiasValorizadas
                             {
                                 C5Filial = SC50.C5Filial,
                                 C5Num = SC50.C5Num,
                                 UAUnumage = a.UaUnumage,
                                 C5XNMMed = SC50.C5XNmmed,
                                 C5XDtCir = $"{SC50.C5XDtcir.Substring(6,2)}/{SC50.C5XDtcir.Substring(4, 2)}/{SC50.C5XDtcir.Substring(0, 4)}",
                                 C6Upatrim = SC60.C6Upatrim,
                                 C5UProcess = SC50.C5Uproces,
                                 C5XNmPac = SC50.C5XNmpac,
                                 C5XNmPla = SC50.C5XNmpla,
                                 TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC50.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                                 Patrimonio = c.Pa1Despat ?? "SEM PATRIMONIO",
                                 C6Produto = SC60.C6Produto,
                                 C5NomClie = SC50.C5Nomclie,
                                 B1Desc = SB10.B1Desc,
                                 C5NomVend = SC50.C5Nomvend,
                                 C6QtdVen = SC60.C6Qtdven,
                                 C6PrcVen = SC60.C6Prcven,
                                 C6QtdEnt = SC60.C6Qtdent,
                                 C6Valor = SC60.C6Valor,
                                 Total = SC60.C6Qtdent * SC60.C6Prcven,
                                 A3XLogin = SA30.A3Xlogin,
                                 Faturado = SC50.C5Nota != "" || (SC50.C5Liberok == "E" && SC50.C5Blq == "") ? "S" : "N"
                             }
                             ).ToList();

                Relatorio.ForEach(x =>
                {
                    Pedido += x.C6QtdVen;
                    QtdFaturada += x.C6QtdEnt;
                    VlUnitario += x.C6PrcVen;
                    TotalPedido += x.C6Valor;
                    TotalFat += x.Total;
                });

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasValorizadas",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {

                Inicio = DataInicio;
                Fim = DataFim;
                Pedido = 0;
                QtdFaturada = 0;
                VlUnitario = 0;
                TotalPedido = 0;
                TotalFat = 0;

                string user = User.Identity.Name.Split("@")[0].ToUpper();

                Relatorio = (from SC50 in Protheus.Sc5010s
                             join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                             join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             join PA10 in Protheus.Pa1010s on new { Filial = SC60.C6Filial, Patrim = SC60.C6Upatrim } equals new { Filial = PA10.Pa1Filial, Patrim = PA10.Pa1Codigo } into Sr
                             from c in Sr.DefaultIfEmpty()
                             join SUA10 in Protheus.Sua010s on new { Filial = SC50.C5Filial, Proces = SC50.C5Uproces } equals new { Filial = SUA10.UaFilial, Proces = SUA10.UaNum } into Rs
                             from a in Rs.DefaultIfEmpty()
                             where SC50.DELET != "*" && SC60.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*" && c.DELET != "*"
                             && a.DELET != "*" && SC50.C5Liberok != "E" && (int)(object)SC50.C5XDtcir >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SC50.C5XDtcir <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             && (SA30.A3Xlogin == user || SA30.A3Xlogsup == user) && SC50.C5Utpoper == "F" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                             orderby SC50.C5XDtcir descending
                             select new RelatorioCirurgiasValorizadas
                             {
                                 C5Filial = SC50.C5Filial,
                                 C5Num = SC50.C5Num,
                                 UAUnumage = a.UaUnumage,
                                 C5XNMMed = SC50.C5XNmmed,
                                 C5XDtCir = $"{SC50.C5XDtcir.Substring(6, 2)}/{SC50.C5XDtcir.Substring(4, 2)}/{SC50.C5XDtcir.Substring(0, 4)}",
                                 C6Upatrim = SC60.C6Upatrim,
                                 C5UProcess = SC50.C5Uproces,
                                 C5XNmPac = SC50.C5XNmpac,
                                 C5XNmPla = SC50.C5XNmpla,
                                 TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC50.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : "OUTROS",
                                 Patrimonio = c.Pa1Despat ?? "SEM PATRIMONIO",
                                 C6Produto = SC60.C6Produto,
                                 C5NomClie = SC50.C5Nomclie,
                                 B1Desc = SB10.B1Desc,
                                 C5NomVend = SC50.C5Nomvend,
                                 C6QtdVen = SC60.C6Qtdven,
                                 C6PrcVen = SC60.C6Prcven,
                                 C6QtdEnt = SC60.C6Qtdent,
                                 C6Valor = SC60.C6Valor,
                                 Total = SC60.C6Qtdent * SC60.C6Prcven,
                                 A3XLogin = SA30.A3Xlogin,
                                 Faturado = SC50.C5Nota != "" || (SC50.C5Liberok == "E" && SC50.C5Blq == "") ? "S" : "N"
                             }
                             ).ToList();

                Relatorio.ForEach(x =>
                {
                    Pedido += x.C6QtdVen;
                    QtdFaturada += x.C6QtdEnt;
                    VlUnitario += x.C6PrcVen;
                    TotalPedido += x.C6Valor;
                    TotalFat += x.Total;
                });

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Cirurgias Valorizadas");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Cirurgias Valorizadas");

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
                sheet.Cells[1, 13].Value = "PRODUTO";
                sheet.Cells[1, 14].Value = "DESC. PRODUTO";
                sheet.Cells[1, 15].Value = "QTD PEDIDO";
                sheet.Cells[1, 16].Value = "QTD FATURADA";
                sheet.Cells[1, 17].Value = "VL. UNITARIO";
                sheet.Cells[1, 18].Value = "TOTAL PEDIDO";
                sheet.Cells[1, 19].Value = "TOTAL FAT.";
                sheet.Cells[1, 20].Value = "FATURADO?";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.C5XDtCir;
                    sheet.Cells[i, 2].Value = Pedido.C5Filial;
                    sheet.Cells[i, 3].Value = Pedido.C5Num;
                    sheet.Cells[i, 4].Value = Pedido.UAUnumage;
                    sheet.Cells[i, 5].Value = Pedido.TipoOper;
                    sheet.Cells[i, 6].Value = Pedido.C5NomVend;
                    sheet.Cells[i, 7].Value = Pedido.C5XNMMed;
                    sheet.Cells[i, 8].Value = Pedido.C5XNmPac;
                    sheet.Cells[i, 9].Value = Pedido.C5NomClie;
                    sheet.Cells[i, 10].Value = Pedido.C5XNmPla;
                    sheet.Cells[i, 11].Value = Pedido.C6Upatrim;
                    sheet.Cells[i, 12].Value = Pedido.Patrimonio;
                    sheet.Cells[i, 13].Value = Pedido.C6Produto;
                    sheet.Cells[i, 14].Value = Pedido.B1Desc;
                    sheet.Cells[i, 15].Value = Pedido.C6QtdVen;
                    sheet.Cells[i, 16].Value = Pedido.C6QtdEnt;
                    sheet.Cells[i, 17].Value = Pedido.C6PrcVen;
                    sheet.Cells[i, 17].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 18].Value = Pedido.C6Valor;
                    sheet.Cells[i, 18].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 19].Value = Pedido.Total;
                    sheet.Cells[i, 19].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 20].Value = Pedido.Faturado;

                    i++;
                });


                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CirurgiasValorizadas.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasValorizadas Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
