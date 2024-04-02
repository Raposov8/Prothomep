using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class InfCirurgiaInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioInfCirurgia> Relatorio { get; set; } = new List<RelatorioInfCirurgia>();

        public InfCirurgiaInterModel(TOTVSINTERContext context,ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        public string Nmage { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string Numage)
        {
            try
            {
                Nmage = Numage;

                Relatorio = (from PAC10 in Protheus.Pac010s
                             join SA10 in Protheus.Sa1010s on new { Cod = PAC10.PacClient, Loj = PAC10.PacLojent } equals new { Cod = SA10.A1Cod, Loj = SA10.A1Loja }
                             join SC50 in Protheus.Sc5010s on new { Filial = PAC10.PacFilial, Numpro = PAC10.PacNumpro } equals new { Filial = SC50.C5Filial, Numpro = SC50.C5Uproces }
                             where PAC10.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC50.DELET != "*" && PAC10.PacNumage == Numage
                             && (SC50.C5Utpoper != "T" && SC50.C5Utpoper != "F")
                             select new RelatorioInfCirurgia
                             {
                                 PacNmPac = PAC10.PacNmpac,
                                 PacDtCir = PAC10.PacDtcir,
                                 PacClient = PAC10.PacClient,
                                 PacLojent = PAC10.PacLojent,
                                 A1Nome = SA10.A1Nome,
                                 C5Num = SC50.C5Num,
                                 C5Utpoper = SC50.C5Utpoper,
                                 C5Nota = SC50.C5Nota
                             }
                             ).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "InfCirurgiaInter",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {

                Relatorio = (from PAC10 in Protheus.Pac010s
                             join SA10 in Protheus.Sa1010s on new { Cod = PAC10.PacClient, Loj = PAC10.PacLojent } equals new { Cod = SA10.A1Cod, Loj = SA10.A1Loja }
                             join SC50 in Protheus.Sc5010s on new { Filial = PAC10.PacFilial, Numpro = PAC10.PacNumpro } equals new { Filial = SC50.C5Filial, Numpro = SC50.C5Uproces }
                             where PAC10.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC50.DELET != "*" && PAC10.PacNumage == Nmage
                             && (SC50.C5Utpoper != "T" && SC50.C5Utpoper != "F")
                             select new RelatorioInfCirurgia
                             {
                                 PacNmPac = PAC10.PacNmpac,
                                 PacDtCir = PAC10.PacDtcir,
                                 PacClient = PAC10.PacClient,
                                 PacLojent = PAC10.PacLojent,
                                 A1Nome = SA10.A1Nome,
                                 C5Num = SC50.C5Num,
                                 C5Utpoper = SC50.C5Utpoper,
                                 C5Nota = SC50.C5Nota
                             }
                             ).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Inf Cirurgia Inter");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Inf Cirurgia Inter");


                sheet.Cells[1, 1].Value = "PAC_NMPAC";
                sheet.Cells[1, 2].Value = "PAC_DTCIR";
                sheet.Cells[1, 3].Value = "PAC_CLIENT";
                sheet.Cells[1, 4].Value = "PAC_LOJENT";
                sheet.Cells[1, 5].Value = "A1_NOME";
                sheet.Cells[1, 6].Value = "C5_NUM";
                sheet.Cells[1, 7].Value = "C5_UTPOPER";
                sheet.Cells[1, 8].Value = "C5_NOTA";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.PacNmPac;
                    sheet.Cells[i, 2].Value = Pedido.PacDtCir;
                    sheet.Cells[i, 3].Value = Pedido.PacClient;
                    sheet.Cells[i, 4].Value = Pedido.PacLojent;
                    sheet.Cells[i, 5].Value = Pedido.A1Nome;
                    sheet.Cells[i, 6].Value = Pedido.C5Num;
                    sheet.Cells[i, 7].Value = Pedido.C5Utpoper;
                    sheet.Cells[i, 8].Value = Pedido.C5Nota;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "InfCirurgiaInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "InfCirurgiaInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
