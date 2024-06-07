using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;
using SGID.Models.Gerencial;

namespace SGID.Pages.Relatorios.Gerencial
{
    [Authorize]
    public class TabelasModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioTabela> Relatorios { get; set; } = new List<RelatorioTabela>();

        public TabelasModel(TOTVSDENUOContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public int CurrentPage { get; set; }
        public int TotalPages { get; set; }

        public void OnGet(int Id)
        {
            try
            {
                var Data = DateTime.Now.ToString("yyyy/MM/dd").Replace("/", "");

                TotalPages = (int)Math.Ceiling(decimal.Divide((from SB10 in Protheus.Sb1010s
                                                               join SBM10 in Protheus.Sbm010s on SB10.B1Grupo equals SBM10.BmGrupo
                                                               join DA10 in Protheus.Da1010s on SB10.B1Cod equals DA10.Da1Codpro
                                                               join DA00 in Protheus.Da0010s on DA10.Da1Codtab equals DA00.Da0Codtab
                                                               where SB10.DELET != "*" && SBM10.DELET != "*" && DA10.DELET != "*" && DA00.DELET != "*"
                                                               select new
                                                               {
                                                                   CodProd = SB10.B1Cod,
                                                                   DescProd = SB10.B1Desc,
                                                                   DescGrup = SBM10.BmDesc,
                                                                   Bloqueado = SB10.B1Msblql != "1" ? "N" : "S",
                                                                   Fabricante = SB10.B1Fabric,
                                                                   CodTab = DA10.Da1Prcven
                                                               }
                             ).Count(), 20));

                Relatorios = (from SB10 in Protheus.Sb1010s
                              join SBM10 in Protheus.Sbm010s on SB10.B1Grupo equals SBM10.BmGrupo
                              join DA10 in Protheus.Da1010s on SB10.B1Cod equals DA10.Da1Codpro
                              join DA00 in Protheus.Da0010s on DA10.Da1Codtab equals DA00.Da0Codtab
                              where SB10.DELET != "*" && SBM10.DELET != "*" && DA10.DELET != "*" && DA00.DELET != "*"
                              select new RelatorioTabela
                              {
                                  CodProd = SB10.B1Cod,
                                  DescProd = SB10.B1Desc,
                                  DescGrup = SBM10.BmDesc,
                                  Bloqueado = SB10.B1Msblql != "1" ? "N" : "S",
                                  Fabricante = SB10.B1Fabric,
                                  CodTab = DA10.Da1Prcven
                              }
                             ).Skip((Id - 1) * 20).Take(20).ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Tabelas",user);
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {
                #region Consultas
                var Data = DateTime.Now.ToString("yyyy/MM/dd").Replace("/", "");

                Relatorios = (from SB10 in Protheus.Sb1010s
                              join SBM10 in Protheus.Sbm010s on SB10.B1Grupo equals SBM10.BmGrupo
                              join DA10 in Protheus.Da1010s on SB10.B1Cod equals DA10.Da1Codpro
                              join DA00 in Protheus.Da0010s on DA10.Da1Codtab equals DA00.Da0Codtab
                              where SB10.DELET != "*" && SBM10.DELET != "*" && DA10.DELET != "*" && DA00.DELET != "*"
                              select new RelatorioTabela
                              {
                                  CodProd = SB10.B1Cod,
                                  DescProd = SB10.B1Desc,
                                  DescGrup = SBM10.BmDesc,
                                  Bloqueado = SB10.B1Msblql != "1" ? "N" : "S",
                                  Fabricante = SB10.B1Fabric,
                                  CodTab = DA10.Da1Prcven
                              }
                             ).ToList();
                #endregion

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Tabela");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Tabela");

                sheet.Cells[1, 1].Value = "Cod.Prod";
                sheet.Cells[1, 2].Value = "Desc.Produto";
                sheet.Cells[1, 3].Value = "Desc.Grupo";
                sheet.Cells[1, 4].Value = "Bloqueado";
                sheet.Cells[1, 5].Value = "Fabricante";
                sheet.Cells[1, 6].Value = "PrcVend";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.CodProd;
                    sheet.Cells[i, 2].Value = Pedido.DescProd;
                    sheet.Cells[i, 3].Value = Pedido.DescGrup;
                    sheet.Cells[i, 4].Value = Pedido.Bloqueado;
                    sheet.Cells[i, 5].Value = Pedido.Fabricante;
                    sheet.Cells[i, 6].Value = Pedido.CodTab;
                    i++;
                });


                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Tabela.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Tabela Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
