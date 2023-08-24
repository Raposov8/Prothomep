using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class DemonstracaoInterModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioConserto> Relatorio { get; set; } = new List<RelatorioConserto>();

        public DemonstracaoInterModel( ApplicationDbContext sgid)
        {
            SGID = sgid;
        }
        public void OnGet()
        {
            try
            {
                Relatorio = new List<RelatorioConserto>(); /*DB.NfdemoInters.Where(x => x.Tipo == "D")
                    .Select(x => new RelatorioConserto
                    {

                        Filial = x.Filial,
                        Nf = x.Nf,
                        Serie = x.Serie,
                        Emissao = x.Emissao,
                        CodCli = x.Codcli,
                        Loja = x.Loja,
                        Cliente = x.Cliente,
                        Dias = (int)(DateTime.Now - x.Emissao).Value.TotalDays

                    }).ToList();

                Relatorio = Relatorio.OrderByDescending(x => x.Dias).ToList();*/
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "DemonstracaoInter",user);
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {

                Relatorio = new List<RelatorioConserto>();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Demonstracao Inter");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Demonstracao Inter");


                sheet.Cells[1, 1].Value = "DIAS";
                sheet.Cells[1, 2].Value = "FILIAL";
                sheet.Cells[1, 3].Value = "NF";
                sheet.Cells[1, 4].Value = "SERIE";
                sheet.Cells[1, 5].Value = "EMISSAO";
                sheet.Cells[1, 6].Value = "COD.CLIENTE";
                sheet.Cells[1, 7].Value = "LOJA";
                sheet.Cells[1, 8].Value = "CLIENTE";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Dias;
                    sheet.Cells[i, 2].Value = Pedido.Filial;
                    sheet.Cells[i, 3].Value = Pedido.Nf;
                    sheet.Cells[i, 4].Value = Pedido.Serie;
                    sheet.Cells[i, 5].Value = Pedido.Emissao;
                    sheet.Cells[i, 6].Value = Pedido.CodCli;
                    sheet.Cells[i, 7].Value = Pedido.Loja;
                    sheet.Cells[i, 8].Value = Pedido.Cliente;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DemonstracaoInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "DemonstracaoInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
