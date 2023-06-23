using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;
using SGID.Models.Qualidade;

namespace SGID.Pages.Relatorios.Qualidade
{
    [Authorize]
    public class RegistroxProdutosModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public List<RelatorioRegistroxProdutos> Relatorios { get; set; } = new List<RelatorioRegistroxProdutos>();

        public RegistroxProdutosModel(TOTVSDENUOContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public string Registro { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string Registro)
        {
            try
            {
                this.Registro = Registro;

                Relatorios = Protheus.Sb1010s.Where(x => x.DELET != "*" && x.B1Reganvi == Registro)
                    .Select(x => new RelatorioRegistroxProdutos
                    {
                        CodProd = x.B1Cod,
                        Descricao = x.B1Desc,
                        RegAnvisa = x.B1Reganvi,
                        ValidRegis = $"{x.B1Dtvalid.Substring(6, 2)}/{x.B1Dtvalid.Substring(4, 2)}/{x.B1Dtvalid.Substring(0, 4)}",
                        ValidAntes = $"{x.B1Exdtval.Substring(6, 2)}/{x.B1Exdtval.Substring(4, 2)}/{x.B1Exdtval.Substring(0, 4)}",
                        Vigente = x.B1Vigente.ToUpper(),
                        Bloqueado = x.B1Msblql == "1" ? "S" : "N"
                    }).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "RegistorxProdutos",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(string Registro)
        {
            try
            {
                this.Registro = Registro;

                Relatorios = Protheus.Sb1010s.Where(x => x.DELET != "*" && x.B1Reganvi == Registro)
                    .Select(x => new RelatorioRegistroxProdutos
                    {
                        CodProd = x.B1Cod,
                        Descricao = x.B1Desc,
                        RegAnvisa = x.B1Reganvi,
                        ValidRegis = $"{x.B1Dtvalid.Substring(6, 2)}/{x.B1Dtvalid.Substring(4, 2)}/{x.B1Dtvalid.Substring(0, 4)}",
                        ValidAntes = $"{x.B1Exdtval.Substring(6, 2)}/{x.B1Exdtval.Substring(4, 2)}/{x.B1Exdtval.Substring(0, 4)}",
                        Vigente = x.B1Vigente.ToUpper(),
                        Bloqueado = x.B1Msblql == "1" ? "S" : "N"
                    }).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("RegistroxProdutos");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "RegistroxProdutos");

                sheet.Cells[1, 1].Value = "Cod. Produto";
                sheet.Cells[1, 2].Value = "Desc. Produto";
                sheet.Cells[1, 3].Value = "Reg. ANVISA";
                sheet.Cells[1, 4].Value = "Valid. Registro";
                sheet.Cells[1, 5].Value = "Valid. Anterior";
                sheet.Cells[1, 6].Value = "Vigente";
                sheet.Cells[1, 7].Value = "Bloqueado";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.CodProd;
                    sheet.Cells[i, 2].Value = Pedido.Descricao;
                    sheet.Cells[i, 3].Value = Pedido.RegAnvisa;
                    sheet.Cells[i, 4].Value = Pedido.ValidRegis;
                    sheet.Cells[i, 5].Value = Pedido.ValidAntes;
                    sheet.Cells[i, 6].Value = Pedido.Vigente;
                    sheet.Cells[i, 7].Value = Pedido.Bloqueado;

                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "RegistroxProdutos.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "RegistroxProdutos Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
