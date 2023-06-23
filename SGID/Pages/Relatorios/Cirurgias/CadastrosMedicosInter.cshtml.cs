using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Cadastro;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.Cirurgias
{
    [Authorize]
    public class CadastrosMedicosInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioCadastroMedicos> Relatorio { get; set; } = new List<RelatorioCadastroMedicos>();

        public int CurrentPage { get; set; }

        public int TotalPages { get; set; }

        public CadastrosMedicosInterModel(TOTVSINTERContext context,ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        public void OnGet(int id)
        {
            try
            {
                CurrentPage = id;

                TotalPages = (int)Math.Ceiling(decimal.Divide((from SA10 in Protheus.Sa1010s
                                                               join SA30 in Protheus.Sa3010s on SA10.A1Vend equals SA30.A3Cod
                                                               where SA10.DELET != "*" && SA10.A1Msblql != "1" && SA10.A1Clinter == "M"
                                                               && SA30.DELET != "*"
                                                               select SA10.A1Nome).Count(), 20));

                Relatorio = (from SA10 in Protheus.Sa1010s
                             join SA30 in Protheus.Sa3010s on SA10.A1Vend equals SA30.A3Cod
                             where SA10.DELET != "*" && SA10.A1Msblql != "1" && SA10.A1Clinter == "M"
                             && SA30.DELET != "*"
                             orderby SA10.A1Cod descending
                             select new RelatorioCadastroMedicos
                             {
                                 Codigo = SA10.A1Cod,
                                 Nome = SA10.A1Nome,
                                 CodVenda = SA10.A1Vend,
                                 Nome2 = SA30.A3Nome,
                                 Telefone = SA10.A1Tel,
                                 Celular = SA10.A1Celular,
                                 Municipio = SA10.A1Mun,
                                 Estado = SA10.A1Est
                             }
                            ).Skip((CurrentPage - 1) * 20).Take(20).ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CadastroMedicosInter",user);

            }
        }
        public IActionResult OnPostExport()
        {
            try
            {
                Relatorio = (from SA10 in Protheus.Sa1010s
                             join SA30 in Protheus.Sa3010s on SA10.A1Vend equals SA30.A3Cod
                             where SA10.DELET != "*" && SA10.A1Msblql != "1" && SA10.A1Clinter == "M"
                             && SA30.DELET != "*"
                             orderby SA10.A1Cod descending
                             select new RelatorioCadastroMedicos
                             {
                                 Codigo = SA10.A1Cod,
                                 Nome = SA10.A1Nome,
                                 CodVenda = SA10.A1Vend,
                                 Nome2 = SA30.A3Nome,
                                 Telefone = SA10.A1Tel,
                                 Celular = SA10.A1Celular,
                                 Municipio = SA10.A1Mun,
                                 Estado = SA10.A1Est
                             }
                            ).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Cadastros Medicos");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Cadastros Medicos");

                sheet.Cells[1, 1].Value = "Codigo";
                sheet.Cells[1, 2].Value = "Nome";
                sheet.Cells[1, 3].Value = "CodVenda";
                sheet.Cells[1, 4].Value = "Nome2";
                sheet.Cells[1, 5].Value = "Telefone";
                sheet.Cells[1, 6].Value = "Celular";
                sheet.Cells[1, 7].Value = "Municipio";
                sheet.Cells[1, 8].Value = "Estado";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Codigo;
                    sheet.Cells[i, 2].Value = Pedido.Nome;
                    sheet.Cells[i, 3].Value = Pedido.CodVenda;
                    sheet.Cells[i, 4].Value = Pedido.Nome2;
                    sheet.Cells[i, 5].Value = Pedido.Telefone;
                    sheet.Cells[i, 6].Value = Pedido.Celular;
                    sheet.Cells[i, 7].Value = Pedido.Municipio;
                    sheet.Cells[i, 8].Value = Pedido.Estado;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CadastroMedicosInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CadastroMedicosInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
