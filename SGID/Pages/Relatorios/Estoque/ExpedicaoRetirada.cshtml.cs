using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Denuo;
using SGID.Models.Estoque;

namespace SGID.Pages.Relatorios.Estoque
{

    [Authorize]
    public class ExpedicaoRetiradaModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioExpedicaoRetirada> Relatorio { get; set; } = new List<RelatorioExpedicaoRetirada>();

        public ExpedicaoRetiradaModel(TOTVSDENUOContext context,ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }
        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(DateTime DataInicio,DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                Relatorio = (from PAC10 in Protheus.Pac010s
                             join SA10 in Protheus.Sa1010s on new { Cod = PAC10.PacClient, Loja = PAC10.PacLojent } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join DA40 in Protheus.Da4010s on PAC10.PacMotent equals DA40.Da4Cod into Sr
                             from P1 in Sr.DefaultIfEmpty()
                             join DA401 in Protheus.Da4010s on PAC10.PacMotent equals DA401.Da4Cod into Pr
                             from P2 in Pr.DefaultIfEmpty()
                             where PAC10.DELET != "*" && P2.DELET != "*" && P1.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1"
                             && (int)(object)PAC10.PacDtcir >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)PAC10.PacDtcir <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             select new RelatorioExpedicaoRetirada
                             {
                                 Agendamento = PAC10.PacNumage,
                                 AgPrincipal = PAC10.PacXagepr,
                                 CodCliente = PAC10.PacClient,
                                 LojaCliente = PAC10.PacLojent,
                                 Nome = SA10.A1Nome,
                                 TipoOper = PAC10.PacTpoper,
                                 DTCirurgia = $"{PAC10.PacDtcir.Substring(6, 2)}/{PAC10.PacDtcir.Substring(4, 2)}/{PAC10.PacDtcir.Substring(0, 4)}",
                                 HoraCirurgia = PAC10.PacHrciru,
                                 Status = PAC10.PacStatus == "1" ? "PENDENTE" : PAC10.PacStatus == "2" ? "LIBERADO" : PAC10.PacStatus == "3" ? "PROCESSADO" : PAC10.PacStatus == "4" ? "CANCELADO" : "ND",
                                 CodCanc = PAC10.PacCodcan == "1" ? "NÃO AUTORIZADO" : PAC10.PacCodcan == "2" ? "CIRURGIA CANCELADA" : PAC10.PacCodcan == "3" ? "ALTERAÇÃO PEDIDO" : PAC10.PacCodcan == "4" ? "FALTA MATERIAL" : PAC10.PacCodcan == "9" ? "OUTROS" : PAC10.PacCodcan == "" ? "" : "ND",
                                 MotivoCanc = PAC10.PacMotcan,
                                 DTExpedicao = $"{PAC10.PacDtexpe.Substring(6, 2)}/{PAC10.PacDtexpe.Substring(4, 2)}/{PAC10.PacDtexpe.Substring(0, 4)}",
                                 HRExpedicao = PAC10.PacHrexpe,
                                 MotEntrega = P1.Da4Nome,
                                 DTRetirada = $"{PAC10.PacDtreti.Substring(6, 2)}/{PAC10.PacDtreti.Substring(4, 2)}/{PAC10.PacDtreti.Substring(0, 4)}",
                                 HRRetirada = PAC10.PacHrreti,
                                 MotRetirada = P2.Da4Nome
                             }
                             ).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ExpedicaoRetirada",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {

                Relatorio = (from PAC10 in Protheus.Pac010s
                             join SA10 in Protheus.Sa1010s on new { Cod = PAC10.PacClient, Loja = PAC10.PacLojent } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join DA40 in Protheus.Da4010s on PAC10.PacMotent equals DA40.Da4Cod into Sr
                             from P1 in Sr.DefaultIfEmpty()
                             join DA401 in Protheus.Da4010s on PAC10.PacMotent equals DA401.Da4Cod into Pr
                             from P2 in Pr.DefaultIfEmpty()
                             where PAC10.DELET != "*" && P2.DELET != "*" && P1.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1"
                             && (int)(object)PAC10.PacDtcir >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)PAC10.PacDtcir <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             select new RelatorioExpedicaoRetirada
                             {
                                 Agendamento = PAC10.PacNumage,
                                 AgPrincipal = PAC10.PacXagepr,
                                 CodCliente = PAC10.PacClient,
                                 LojaCliente = PAC10.PacLojent,
                                 Nome = SA10.A1Nome,
                                 TipoOper = PAC10.PacTpoper,
                                 DTCirurgia = $"{PAC10.PacDtcir.Substring(6, 2)}/{PAC10.PacDtcir.Substring(4, 2)}/{PAC10.PacDtcir.Substring(0, 4)}",
                                 HoraCirurgia = PAC10.PacHrciru,
                                 Status = PAC10.PacStatus == "1" ? "PENDENTE" : PAC10.PacStatus == "2" ? "LIBERADO" : PAC10.PacStatus == "3" ? "PROCESSADO" : PAC10.PacStatus == "4" ? "CANCELADO" : "ND",
                                 CodCanc = PAC10.PacCodcan == "1" ? "NÃO AUTORIZADO" : PAC10.PacCodcan == "2" ? "CIRURGIA CANCELADA" : PAC10.PacCodcan == "3" ? "ALTERAÇÃO PEDIDO" : PAC10.PacCodcan == "4" ? "FALTA MATERIAL" : PAC10.PacCodcan == "9" ? "OUTROS" : PAC10.PacCodcan == "" ? "" : "ND",
                                 MotivoCanc = PAC10.PacMotcan,
                                 DTExpedicao = $"{PAC10.PacDtexpe.Substring(6, 2)}/{PAC10.PacDtexpe.Substring(4, 2)}/{PAC10.PacDtexpe.Substring(0, 4)}",
                                 HRExpedicao = PAC10.PacHrexpe,
                                 MotEntrega = P1.Da4Nome,
                                 DTRetirada = $"{PAC10.PacDtreti.Substring(6, 2)}/{PAC10.PacDtreti.Substring(4, 2)}/{PAC10.PacDtreti.Substring(0, 4)}",
                                 HRRetirada = PAC10.PacHrreti,
                                 MotRetirada = P2.Da4Nome
                             }
                             ).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("ExpedicaoRetirada");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "ExpedicaoRetirada");

                sheet.Cells[1, 1].Value = "Agendamento";
                sheet.Cells[1, 2].Value = "Ag.Principal";
                sheet.Cells[1, 3].Value = "Cod.Cliente";
                sheet.Cells[1, 4].Value = "Loja Cliente";
                sheet.Cells[1, 5].Value = "Nome";
                sheet.Cells[1, 6].Value = "Tipo Oper.";
                sheet.Cells[1, 7].Value = "Data Cirurgia";
                sheet.Cells[1, 8].Value = "Hora Cirurgia";
                sheet.Cells[1, 9].Value = "Status";
                sheet.Cells[1, 10].Value = "Cod.Canc.";
                sheet.Cells[1, 11].Value = "Motivo Canc.";
                sheet.Cells[1, 12].Value = "DT.Expedição";
                sheet.Cells[1, 13].Value = "HR Expedição";
                sheet.Cells[1, 14].Value = "Mot.Entrega";
                sheet.Cells[1, 15].Value = "DT Retirada";
                sheet.Cells[1, 16].Value = "HR Retirada";
                sheet.Cells[1, 17].Value = "Mot.Retirada";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Agendamento;
                    sheet.Cells[i, 2].Value = Pedido.AgPrincipal;
                    sheet.Cells[i, 3].Value = Pedido.CodCliente;
                    sheet.Cells[i, 4].Value = Pedido.LojaCliente;
                    sheet.Cells[i, 5].Value = Pedido.Nome;
                    sheet.Cells[i, 6].Value = Pedido.TipoOper;
                    sheet.Cells[i, 7].Value = Pedido.DTCirurgia;
                    sheet.Cells[i, 8].Value = Pedido.HoraCirurgia;
                    sheet.Cells[i, 9].Value = Pedido.Status;
                    sheet.Cells[i, 10].Value = Pedido.CodCanc;
                    sheet.Cells[i, 11].Value = Pedido.MotivoCanc;
                    sheet.Cells[i, 12].Value = Pedido.DTExpedicao;
                    sheet.Cells[i, 13].Value = Pedido.HRExpedicao;
                    sheet.Cells[i, 14].Value = Pedido.MotEntrega;
                    sheet.Cells[i, 15].Value = Pedido.DTRetirada;
                    sheet.Cells[i, 16].Value = Pedido.HRRetirada;
                    sheet.Cells[i, 17].Value = Pedido.MotRetirada;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExpedicaoRetirada.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ExpedicaoRetirada Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
