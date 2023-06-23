using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;

namespace SGID.Pages.Utilitarios
{
    public class ListarCRMDenuoModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public int Total { get; set; }

        public List<Sa1010> Medicos { get; set; }


        public ListarCRMDenuoModel(TOTVSDENUOContext denuo,ApplicationDbContext sgid)
        {
            ProtheusDenuo = denuo;
            SGID = sgid;
        }
        public void OnGet()
        {
            Medicos = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && x.A1Vend != "" && x.A1Crm == "").OrderBy(x => x.A1Nvend).ToList();
            Total = Medicos.Count;
        }

        public IActionResult OnPostExport()
        {
            try
            {
                Medicos = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && x.A1Vend != "" && x.A1Crm == "").OrderBy(x=> x.A1Nvend).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Medicos Sem CRM Denuo");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Medicos Sem CRM Denuo");

                sheet.Cells[1, 1].Value = "Codigo";
                sheet.Cells[1, 2].Value = "Nome Medico";
                sheet.Cells[1, 3].Value = "CRM";
                sheet.Cells[1, 4].Value = "Vendedor";

                int i = 2;

                Medicos.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.A1Cod;
                    sheet.Cells[i, 2].Value = Pedido.A1Nome;
                    sheet.Cells[i, 3].Value = Pedido.A1Crm;
                    sheet.Cells[i, 4].Value = Pedido.A1Nvend;

                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MedicosSemCRMDenuo.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ListarCRMDenuo Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
