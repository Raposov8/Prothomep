using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class DemonstracaoInterModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext Protheus { get; set; }

        public List<RelatorioConserto> Relatorio { get; set; } = new List<RelatorioConserto>();

        public DemonstracaoInterModel( ApplicationDbContext sgid, TOTVSINTERContext inter)
        {
            SGID = sgid;
            Protheus = inter;
        }
        public void OnGet()
        {
            try
            {

                Relatorio = (from SD20 in Protheus.Sd2010s
                             join SA10 in Protheus.Sa1010s on new { Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SD10 in Protheus.Sd1010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Cliente = SD10.D1Fornece, Loja = SD10.D1Loja } into Sr
                             from m in Sr.DefaultIfEmpty()
                             where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && (SD20.D2Cf == "5912" || SD20.D2Cf == "6912") && SD20.D2Tipo == "N"
                             && SD20.D2Serie == "2" && m.D1Cf == null
                             select new RelatorioConserto
                             {
                                 Filial = SD20.D2Filial,
                                 Nf = SD20.D2Doc,
                                 Serie = SD20.D2Serie,
                                 Emissao = Convert.ToDateTime($"{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(0, 4)}"),
                                 CodCli = SD20.D2Cliente,
                                 Loja = SD20.D2Loja,
                                 Dias = (int)(DateTime.Now - Convert.ToDateTime($"{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(0, 4)}")).TotalDays,
                                 Cliente = SA10.A1Nome,
                                 CF = m.D1Cf
                             }
                             ).ToList();

                Relatorio = Relatorio.DistinctBy(x => x.Nf).ToList();

                Relatorio = Relatorio.OrderByDescending(x => x.Dias).ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "DemonstracaoInter",user);
            }
        }

        public IActionResult OnPostExport(List<RelatorioConserto> Relatorio)
        {
            try
            {

                var data = Convert.ToInt32(DateTime.Now.ToString("yyyy/MM/dd").Replace("/", ""));


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
