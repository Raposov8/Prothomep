using Intergracoes.VExpenses;
using Intergracoes.VExpenses.Despecas;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;

namespace SGID.Pages.Relatorios.Vexpenses
{
    [Authorize(Roles = "Admin,Financeiro,Diretoria")]
    public class RelatorioDespesasModel : PageModel
    {
        private Expenses VExpenseDespesa { get; set; } = new Expenses();

        public List<Despecas> Despecas { get; set; } = new List<Despecas>();

        private ApplicationDbContext SGID { get; set; }
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public RelatorioDespesasModel(ApplicationDbContext Sgid) => SGID = Sgid;
        public void OnGet()
        {
            
        }

        public IActionResult OnPostAsync(DateTime Inicio,DateTime Fim)
        {
            this.Inicio = Inicio;
            this.Fim = Fim;

            Despecas = VExpenseDespesa.GetDespecas(Inicio, Fim).Result;

            return Page();
        }

        public IActionResult OnPostExport(DateTime Inicio, DateTime Fim)
        {
            try
            {
                var lista = VExpenseDespesa.GetDespecas(Inicio, Fim).Result;

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Despesas VExpenses");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Despesas VExpenses");

                sheet.Cells[1, 1].Value = "Data Despesa";
                sheet.Cells[1, 2].Value = "Informada a Despesa em:";
                sheet.Cells[1, 3].Value = "Alterada a Despesa em:";
                sheet.Cells[1, 4].Value = "Titulo";
                sheet.Cells[1, 5].Value = "Status Relatorio";
                sheet.Cells[1, 6].Value = "Tipo Despesa";
                sheet.Cells[1, 7].Value = "Centro de custo";
                sheet.Cells[1, 8].Value = "Valor";
                sheet.Cells[1, 9].Value = "Moeda";
                sheet.Cells[1, 10].Value = "Nome Funcionario";

                int i = 2;

                lista.ForEach(Despesa =>
                {

                    sheet.Cells[i, 1].Value = Despesa.date.Value.ToString("dd/MM/yyyy HH:mm:ss");
                    sheet.Cells[i, 2].Value = Despesa.created_at.Value.ToString("dd/MM/yyyy HH:mm:ss");
                    sheet.Cells[i, 3].Value = Despesa.updated_at.Value.ToString("dd/MM/yyyy HH:mm:ss");
                    sheet.Cells[i, 4].Value = Despesa.title;
                    sheet.Cells[i, 5].Value = Despesa.report.data.status;
                    sheet.Cells[i, 6].Value = Despesa.expense_type.data.description;
                    sheet.Cells[i, 7].Value = Despesa.costs_center.data.name;
                    sheet.Cells[i, 8].Value = Despesa.value;
                    sheet.Cells[i, 8].Style.Numberformat.Format = "0.00";
                    sheet.Cells[i, 9].Value = Despesa.original_currency_iso;
                    sheet.Cells[i, 10].Value = Despesa.user.data.name;

                    i++;
                });


                sheet.Cells[i, 1, i, 6].Merge = true;


                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DespesasVExpense.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Despesa VExpense Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
