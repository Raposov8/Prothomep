using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Gerencial;

namespace SGID.Pages.Relatorios.Gerencial
{
    [Authorize]
    public class TabelasProdNCadastradosInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioTabelasProdNCadastrados> Relatorios { get; set; } = new List<RelatorioTabelasProdNCadastrados>();

        public TabelasProdNCadastradosInterModel(TOTVSINTERContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public string Tabela { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string Tabela)
        {
            try
            {
                this.Tabela = Tabela;

                Relatorios = (from SB10 in Protheus.Sb1010s
                              join SBM10 in Protheus.Sbm010s on SB10.B1Grupo equals SBM10.BmGrupo
                              where SB10.DELET != "*" && SB10.B1Tipo == "PA" && SBM10.DELET != "*"
                              &&
                              !(from DA00 in Protheus.Da0010s
                                join DA10 in Protheus.Da1010s on new { Filial = DA00.Da0Filial, CodTab = DA00.Da0Codtab } equals new { Filial = DA10.Da1Filial, CodTab = DA10.Da1Codtab }
                                where DA00.DELET != "*" && DA10.DELET != "*" && DA10.Da1Codtab == Tabela
                                select DA10.Da1Codpro).ToArray().Contains(SB10.B1Cod)
                              select new RelatorioTabelasProdNCadastrados
                              {
                                  Codigo = SB10.B1Cod,
                                  Descricao = SB10.B1Desc,
                                  Grupo = SBM10.BmDesc,
                                  Bloqueado = SB10.B1Msblql == "1" ? "SIM" : "NÃO",
                                  Fabricante = SB10.B1Fabric
                              }
                             ).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "TabelasProdNCadastradosInter",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {
                #region Consultas
                Relatorios = (from SB10 in Protheus.Sb1010s
                              join SBM10 in Protheus.Sbm010s on SB10.B1Grupo equals SBM10.BmGrupo
                              where SB10.DELET != "*" && SB10.B1Tipo == "PA" && SBM10.DELET != "*"
                              &&
                              !(from DA00 in Protheus.Da0010s
                                join DA10 in Protheus.Da1010s on new { Filial = DA00.Da0Filial, CodTab = DA00.Da0Codtab } equals new { Filial = DA10.Da1Filial, CodTab = DA10.Da1Codtab }
                                where DA00.DELET != "*" && DA10.DELET != "*" && DA10.Da1Codtab == Tabela
                                select DA10.Da1Codpro).ToArray().Contains(SB10.B1Cod)
                              select new RelatorioTabelasProdNCadastrados
                              {
                                  Codigo = SB10.B1Cod,
                                  Descricao = SB10.B1Desc,
                                  Grupo = SBM10.BmDesc,
                                  Bloqueado = SB10.B1Msblql == "1" ? "SIM" : "NÃO",
                                  Fabricante = SB10.B1Fabric
                              }
                              ).ToList();
                #endregion

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("TabelaProdNCadastrados");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "TabelaProdNCadastrados");

                sheet.Cells[1, 1].Value = "Codigo";
                sheet.Cells[1, 2].Value = "Descrição";
                sheet.Cells[1, 3].Value = "Grupo";
                sheet.Cells[1, 4].Value = "Fabricante";
                sheet.Cells[1, 5].Value = "Bloqueado";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Codigo;
                    sheet.Cells[i, 2].Value = Pedido.Descricao;
                    sheet.Cells[i, 3].Value = Pedido.Grupo;
                    sheet.Cells[i, 4].Value = Pedido.Fabricante;
                    sheet.Cells[i, 5].Value = Pedido.Bloqueado;
                    i++;
                });


                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TabelaProdNCadastradosInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "TabelaProdNCadastradosInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
