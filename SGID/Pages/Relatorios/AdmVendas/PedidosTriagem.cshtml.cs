using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.DTO;
using SGID.Models.Denuo;

namespace SGID.Pages.Relatorios.AdmVendas
{

    [Authorize]
    public class PedidosTriagemModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public List<RelatorioPedidosTriagem> Relatorios { get; set; } = new List<RelatorioPedidosTriagem>();

        public PedidosTriagemModel(TOTVSDENUOContext protheus, ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public void OnGet()
        {
            try
            {
                Relatorios = Protheus.Sc5010s.Where(x => x.DELET != "*" && x.C5Nota == "" && x.C5Utpoper == "T")
                    .OrderBy(x => x.C5XDtcir).ThenBy(x => x.C5Emissao).Select(x =>
                    new RelatorioPedidosTriagem
                    {
                        Filial = x.C5Filial,
                        NumPedido = x.C5Num,
                        CodCliente = x.C5Cliente,
                        CodLoja = x.C5Lojacli,
                        NomCliente = x.C5Nomcli.Trim(),
                        DTCirurgia = $"{x.C5XDtcir.Substring(6, 2)}/{x.C5XDtcir.Substring(4, 2)}/{x.C5XDtcir.Substring(0, 4)}",
                        DTEmissao = $"{x.C5Emissao.Substring(6, 2)}/{x.C5Emissao.Substring(4, 2)}/{x.C5Emissao.Substring(0, 4)}",
                        Paciente = x.C5XNmpac.Trim(),
                    }).ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "PedidosTriagem", user);
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {
                Relatorios = Protheus.Sc5010s.Where(x => x.DELET != "*" && x.C5Nota == "" && x.C5Utpoper == "T")
                    .OrderBy(x => x.C5XDtcir).ThenBy(x => x.C5Emissao).Select(x =>
                    new RelatorioPedidosTriagem
                    {
                        Filial = x.C5Filial,
                        NumPedido = x.C5Num,
                        CodCliente = x.C5Cliente,
                        CodLoja = x.C5Lojacli,
                        NomCliente = x.C5Nomcli.Trim(),
                        DTCirurgia = $"{x.C5XDtcir.Substring(6, 2)}/{x.C5XDtcir.Substring(4, 2)}/{x.C5XDtcir.Substring(0, 4)}",
                        DTEmissao = $"{x.C5Emissao.Substring(6, 2)}/{x.C5Emissao.Substring(4, 2)}/{x.C5Emissao.Substring(0, 4)}",
                        Paciente = x.C5XNmpac.Trim(),
                    }).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Pedidos Triagem");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Pedidos Triagem");

                sheet.Cells[1, 1].Value = "Filial";
                sheet.Cells[1, 2].Value = "NumPedido";
                sheet.Cells[1, 3].Value = "CodCliente";
                sheet.Cells[1, 4].Value = "CodLoja";
                sheet.Cells[1, 5].Value = "NomCliente";
                sheet.Cells[1, 6].Value = "DTCirurgia";
                sheet.Cells[1, 7].Value = "DTEmissao";
                sheet.Cells[1, 8].Value = "Paciente";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.NumPedido;
                    sheet.Cells[i, 3].Value = Pedido.CodCliente;
                    sheet.Cells[i, 4].Value = Pedido.CodLoja;
                    sheet.Cells[i, 5].Value = Pedido.NomCliente;
                    sheet.Cells[i, 6].Value = Pedido.DTCirurgia;
                    sheet.Cells[i, 7].Value = Pedido.DTEmissao;
                    sheet.Cells[i, 8].Value = Pedido.Paciente;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PedidosTriagem.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "PedidosTriagem Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
