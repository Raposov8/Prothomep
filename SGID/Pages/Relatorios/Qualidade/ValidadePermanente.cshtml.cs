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
    public class ValidadePermanenteModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioValidadePermanente> Relatorios { get; set; } = new List<RelatorioValidadePermanente>();

        public ValidadePermanenteModel(TOTVSDENUOContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
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

                var query = (from SB80 in Protheus.Sb8010s
                             join SB10 in Protheus.Sb1010s on SB80.B8Produto equals SB10.B1Cod
                             where SB10.DELET != "*" && SB80.DELET != "*" && SB80.B8Saldo > 0
                             && (int)(object)SB80.B8Dtvalid >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SB80.B8Dtvalid <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             select new RelatorioValidadePermanente
                             {
                                 Codigo = SB80.B8Produto,
                                 Descricao = SB10.B1Desc,
                                 Filial = SB80.B8Filial,
                                 Local = SB80.B8Local,
                                 Lote = SB80.B8Lotectl,
                                 LoteFor = SB80.B8Lotefor,
                                 ValidLote = SB80.B8Dtvalid,
                                 Saldo = SB80.B8Saldo

                             }).ToList();


                Relatorios = query.OrderBy(x => x.ValidLote).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ValidadePermanente",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                var query = (from SB80 in Protheus.Sb8010s
                             join SB10 in Protheus.Sb1010s on SB80.B8Produto equals SB10.B1Cod
                             where SB10.DELET != "*" && SB80.DELET != "*" && SB80.B8Saldo > 0
                             && (int)(object)SB80.B8Dtvalid >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SB80.B8Dtvalid <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             select new RelatorioValidadePermanente
                             {
                                 Codigo = SB80.B8Produto,
                                 Descricao = SB10.B1Desc,
                                 Filial = SB80.B8Filial,
                                 Local = SB80.B8Local,
                                 Lote = SB80.B8Lotectl,
                                 LoteFor = SB80.B8Lotefor,
                                 ValidLote = SB80.B8Dtvalid,
                                 Saldo = SB80.B8Saldo

                             }).ToList();


                Relatorios = query.OrderBy(x => x.ValidLote).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("ValidadePermanente");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "ValidadePermanente");

                sheet.Cells[1, 1].Value = "Codigo";
                sheet.Cells[1, 2].Value = "Descrição";
                sheet.Cells[1, 3].Value = "Filial";
                sheet.Cells[1, 4].Value = "Local";
                sheet.Cells[1, 5].Value = "Lote";
                sheet.Cells[1, 6].Value = "Lote For";
                sheet.Cells[1, 7].Value = "Valid Lote";
                sheet.Cells[1, 8].Value = "Saldo";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Codigo;
                    sheet.Cells[i, 2].Value = Pedido.Descricao;
                    sheet.Cells[i, 3].Value = Pedido.Filial;
                    sheet.Cells[i, 4].Value = Pedido.Local;
                    sheet.Cells[i, 5].Value = Pedido.Lote;
                    sheet.Cells[i, 6].Value = Pedido.LoteFor;
                    sheet.Cells[i, 7].Value = $"{Pedido.ValidLote.Substring(6, 2)}/{Pedido.ValidLote.Substring(4, 2)}/{Pedido.ValidLote.Substring(0, 4)}";
                    sheet.Cells[i, 8].Value = Pedido.Saldo;

                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ValidadePermanente.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ValidadePermanente Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
