using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;

namespace SGID.Pages.Cirurgias
{
    [Authorize(Roles = "Admin,GestorVenda,Venda,Diretoria")]
    public class ListarRecursosModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext Protheus { get; set; }
        public List<Pah010> Recursos { get; set; } = new List<Pah010>();
        public ListarRecursosModel(ApplicationDbContext sgid, TOTVSINTERContext protheus) 
        { 
            SGID = sgid;
            Protheus = protheus;
        }
        public void OnGet()
        {
            Recursos = Protheus.Pah010s.Where(x => x.DELET != "*" && x.PahCodins != "000").ToList();
        }

        public IActionResult OnPost(string Tipo,string Status)
        {

            var query = Protheus.Pah010s.Where(x => x.DELET != "*" && x.PahCodins != "000");

            if(!string.IsNullOrEmpty(Tipo) && !string.IsNullOrWhiteSpace(Tipo))
            {
                query = query.Where(x => x.PahTipocontrato == Tipo);
            }

            if (!string.IsNullOrEmpty(Status) && !string.IsNullOrWhiteSpace(Status))
            {
                query = query.Where(x => x.PahMsblql == Status);
            }

            Recursos = query.ToList();

            return Page();
        }

        public IActionResult OnPostExport(string Tipo, string Status)
        {
            try
            {
                var query = Protheus.Pah010s.Where(x => x.DELET != "*" && x.PahCodins != "000");

                if (!string.IsNullOrEmpty(Tipo) && !string.IsNullOrWhiteSpace(Tipo))
                {
                    query = query.Where(x => x.PahTipocontrato == Tipo);
                }

                if (!string.IsNullOrEmpty(Status) && !string.IsNullOrWhiteSpace(Status))
                {
                    query = query.Where(x => x.PahMsblql == Status);
                }

                Recursos = query.ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Recursos");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Recursos");

                sheet.Cells[1, 1].Value = "Nome do Recurso";
                sheet.Cells[1, 2].Value = "Tipo";
                sheet.Cells[1, 3].Value = "Contratação";
                sheet.Cells[1, 4].Value = "Status";

                int i = 2;

                Recursos.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.PahNome;
                    sheet.Cells[i, 2].Value = "Intrumentador";
                    sheet.Cells[i, 3].Value = Pedido.PahTipocontrato == "1"? "CLT": Pedido.PahTipocontrato == "2" ?"PJ":"Não Informado";
                    sheet.Cells[i, 4].Value = Pedido.PahMsblql == "1"? "INATIVO": "ATIVO";

                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Recursos.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Recursos Excel", user);

                return LocalRedirect("/error");
            }
        }

        //Delete
        public JsonResult OnPostExcluir(string Id)
        {
            try
            {
                int id = Convert.ToInt32(Id);
                var Recurso = Protheus.Pah010s.FirstOrDefault(x => x.RECNO == id);

                if (Recurso == null)
                {
                    var Status = new
                    {
                        Status = false
                    };

                    return new JsonResult(Status);
                }
                else
                {
                    var Status = new
                    {
                        Status = true
                    };

                    Recurso.DELET = "*";

                    Protheus.Update(Recurso);
                    Protheus.SaveChanges();

                    return new JsonResult(Status);
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ListarRecursos",user);

                var Status = new
                {
                    Status = false
                };

                return new JsonResult(Status);
            }
        }

    }
}
