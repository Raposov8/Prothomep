using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Denuo;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class RetornoSefazModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioRetornoSefaz> Relatorio = new List<RelatorioRetornoSefaz>();

        public RetornoSefazModel(TOTVSDENUOContext context,ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(DateTime DataInicio,DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                Relatorio = Protheus.Sf3010s.Where(x => x.DELET != "*" && x.F3Codrsef != "100" && x.F3Codrsef != "101" && x.F3Codrsef != "102" && x.F3Codrsef != "155"
                && (int)(object)x.F3Entrada >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)x.F3Entrada <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "") && x.F3Codrsef != "*")
                    .Select(x => new RelatorioRetornoSefaz
                    {
                        Filial = x.F3Filial,
                        DtEntrada = $"{x.F3Entrada.Substring(6, 2)}/{x.F3Entrada.Substring(4, 2)}/{x.F3Entrada.Substring(0, 4)}",
                        DtEmissao = $"{x.F3Emissao.Substring(6, 2)}/{x.F3Emissao.Substring(4, 2)}/{x.F3Emissao.Substring(0, 4)}",
                        NotaFiscal = x.F3Nfiscal,
                        Serie = x.F3Serie,
                        CliFor = x.F3Cliefor,
                        Loja = x.F3Loja,
                        CFOP = x.F3Cfo,
                        CodSefaz = x.F3Codrsef,
                        Observacao = x.F3Observ
                    }).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "RetornoSefaz",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {

                Relatorio = Protheus.Sf3010s.Where(x => x.DELET != "*" && x.F3Codrsef != "100" && x.F3Codrsef != "101" && x.F3Codrsef != "102" && x.F3Codrsef != "155"
                && (int)(object)x.F3Entrada >= (int)(object)Inicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)x.F3Entrada <= (int)(object)Fim.ToString("yyyy/MM/dd").Replace("/", "") && x.F3Codrsef != "*")
                    .Select(x => new RelatorioRetornoSefaz
                    {
                        Filial = x.F3Filial,
                        DtEntrada = $"{x.F3Entrada.Substring(6, 2)}/{x.F3Entrada.Substring(4, 2)}/{x.F3Entrada.Substring(0, 4)}",
                        DtEmissao = $"{x.F3Emissao.Substring(6, 2)}/{x.F3Emissao.Substring(4, 2)}/{x.F3Emissao.Substring(0, 4)}",
                        NotaFiscal = x.F3Nfiscal,
                        Serie = x.F3Serie,
                        CliFor = x.F3Cliefor,
                        Loja = x.F3Loja,
                        CFOP = x.F3Cfo,
                        CodSefaz = x.F3Codrsef,
                        Observacao = x.F3Observ
                    }).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("RetornoSefaz");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "RetornoSefaz");

                sheet.Cells[1, 1].Value = "FILIAL";
                sheet.Cells[1, 2].Value = "COD.SEFAZ";
                sheet.Cells[1, 3].Value = "OBSERVAÇÃO";
                sheet.Cells[1, 4].Value = "DT.ENTRADA";
                sheet.Cells[1, 5].Value = "DT.EMISSAO";
                sheet.Cells[1, 6].Value = "NOTA FISCAL";
                sheet.Cells[1, 7].Value = "SERIE";
                sheet.Cells[1, 8].Value = "CLI/FOR";
                sheet.Cells[1, 9].Value = "LOJA";
                sheet.Cells[1, 10].Value = "CFOP";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.CodSefaz;
                    sheet.Cells[i, 3].Value = Pedido.Observacao;
                    sheet.Cells[i, 4].Value = Pedido.DtEntrada;
                    sheet.Cells[i, 5].Value = Pedido.DtEmissao;
                    sheet.Cells[i, 6].Value = Pedido.NotaFiscal;
                    sheet.Cells[i, 7].Value = Pedido.Serie;
                    sheet.Cells[i, 8].Value = Pedido.CliFor;
                    sheet.Cells[i, 9].Value = Pedido.Loja;
                    sheet.Cells[i, 10].Value = Pedido.CFOP;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "RetornoSefaz.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "RetornoSefaz Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
