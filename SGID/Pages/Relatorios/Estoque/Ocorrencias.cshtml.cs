using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Estoque;

namespace SGID.Pages.Relatorios.Estoque
{
    public class OcorrenciasModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        
        public List<Ocorrencia> Relatorio { get; set; } = new List<Ocorrencia>();

        public List<RelatorioOcorrencia> relatorioOcorrencias { get; set; } = new List<RelatorioOcorrencia>();
        public DateTime Inicio { get; set; } 
        public DateTime Fim { get; set; }



        public OcorrenciasModel(ApplicationDbContext sgid)
        {
            SGID = sgid;
        }

        public void OnGet()
        {

        }

        public IActionResult OnPostAsync(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                Relatorio = SGID.Ocorrencias.Where(x => x.DataOcorrencia >= DataInicio && x.DataOcorrencia <= DataFim).ToList();


                relatorioOcorrencias = Relatorio.GroupBy(x => new
                {
                    x.Medico
                })
                .Select(x => new RelatorioOcorrencia
                {
                    Nome = x.Key.Medico,
                    Quant = x.Count()
                }).OrderByDescending(x=> x.Quant).ToList();


                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Relatorio Ocorrencias", user);
                return LocalRedirect("/error");
            }
        }
        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                Relatorio = SGID.Ocorrencias.Where(x => x.DataOcorrencia >= DataInicio && x.DataOcorrencia <= DataFim).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Ocorrencias");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Ocorrencias");

                sheet.Cells[1, 1].Value = "EMPRESA";
                sheet.Cells[1, 2].Value = "DATA DA INCLUSÃO";
                sheet.Cells[1, 3].Value = "CODIGO DA OCORRENCIA";
                sheet.Cells[1, 4].Value = "AGENDAMENTO";
                sheet.Cells[1, 5].Value = "HOSPITAL";
                sheet.Cells[1, 6].Value = "PACIENTE";
                sheet.Cells[1, 7].Value = "MEDICO";
                sheet.Cells[1, 8].Value = "CODIGO DO PRODUTO";
                sheet.Cells[1, 9].Value = "PRODUTO";
                sheet.Cells[1, 10].Value = "PROBLEMA";
                sheet.Cells[1, 11].Value = "AÇÕES";
                sheet.Cells[1, 12].Value = "VENDEDOR";
                sheet.Cells[1, 13].Value = "MES";
                sheet.Cells[1, 14].Value = "ARMAZEM DE ORIGEM";
 
                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Empresa == "01" ? "INTERMEDIC" : "DENUO";
                    sheet.Cells[i, 2].Value = Pedido.DataCriacao.ToString("dd/MM/yyyy HH:mm");
                    sheet.Cells[i, 3].Value = Pedido.Id;
                    sheet.Cells[i, 4].Value = Pedido.Agendamento;
                    sheet.Cells[i, 5].Value = Pedido.Hospital;
                    sheet.Cells[i, 6].Value = Pedido.Paciente;
                    sheet.Cells[i, 7].Value = Pedido.Medico;
                    sheet.Cells[i, 8].Value = Pedido.Produto;
                    sheet.Cells[i, 9].Value = Pedido.Descricao;
                    sheet.Cells[i, 10].Value = Pedido.Problema;
                    sheet.Cells[i, 11].Value = Pedido.Acao;
                    sheet.Cells[i, 12].Value = Pedido.Vendedor;
                    sheet.Cells[i, 13].Value = Pedido.DataOcorrencia.Value.Month;
                    sheet.Cells[i, 14].Value = Pedido.Armazem;

                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Ocorrencias.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Ocorrencias Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
