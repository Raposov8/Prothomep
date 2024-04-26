using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Controladoria.FaturamentoNF.RegistrosFaturamento;
using SGID.Models.Denuo;
using SGID.Models.Diretoria;

namespace SGID.Pages.Relatorios.Diretoria
{
    public class TitulosVencidosModel : PageModel
    {

        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext Protheus { get; set; }
        public List<TitulosVencidos> Relatorio { get; set; } = new List<TitulosVencidos>();

        public List<string> Clientes { get; set; } = new List<string>();
        public string Cliente { get; set; } = "";

        public TitulosVencidosModel(ApplicationDbContext sgid,TOTVSDENUOContext denuo)
        {
            SGID = sgid;
            Protheus = denuo;
        }

        public void OnGet()
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                DateTime data = DateTime.Now;
                string DataInicio = $"{data.Year}{data.Month.ToString("D2")}{data.Day.ToString("D2")}";

                Relatorio = (from SE10 in Protheus.Se1010s
                             join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                             where SE10.DELET != "*"
                             && (int)(object)SE10.E1Vencrea < (int)(object)DataInicio
                             && SA10.A1Clinter != "G" && (SE10.E1Baixa == "" || SE10.E1Saldo > 0)
                             && SE10.E1Tipo != "RA"
                             select new TitulosVencidos
                             {
                                 CodCliente = SE10.E1Cliente,
                                 Cliente = SA10.A1Nome,
                                 NF = SE10.E1Num,
                                 Valor = SE10.E1Valor,
                                 ValorSaldo = SE10.E1Saldo,
                                 DataEmissao = $"{SE10.E1Emissao.Substring(6, 2)}/{SE10.E1Emissao.Substring(4, 2)}/{SE10.E1Emissao.Substring(0, 4)}",
                                 DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}"
                             }
                                   )
                                   .OrderByDescending(x => x.Valor).ToList();

                Clientes = (from SE10 in Protheus.Se1010s
                            join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                            where SE10.DELET != "*"
                            && (int)(object)SE10.E1Vencrea < (int)(object)DataInicio
                            && SA10.A1Clinter != "G" && (SE10.E1Baixa == "" || SE10.E1Saldo > 0)
                            && SE10.E1Tipo != "RA"
                            select new TitulosVencidos
                            {
                                CodCliente = SE10.E1Cliente,
                                Cliente = SA10.A1Nome,
                                Valor = SE10.E1Valor,
                                DataEmissao = $"{SE10.E1Emissao.Substring(6, 2)}/{SE10.E1Emissao.Substring(4, 2)}/{SE10.E1Emissao.Substring(0, 4)}",
                                DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}"
                            }).OrderBy(x => x.Cliente).Select(x => x.Cliente).Distinct().ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "TitulosVencidos", user);
            }
        }

        public IActionResult OnPost(string NReduz)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                DateTime data = DateTime.Now;
                string DataInicio = $"{data.Year}{data.Month.ToString("D2")}{data.Day.ToString("D2")}";

                var query = (from SE10 in Protheus.Se1010s
                             join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                             where SE10.DELET != "*"
                             && (int)(object)SE10.E1Vencrea < (int)(object)DataInicio
                             && SA10.A1Clinter != "G" && (SE10.E1Baixa == "" || SE10.E1Saldo > 0)
                             && SE10.E1Tipo != "RA"
                             select new TitulosVencidos
                             {
                                 CodCliente = SE10.E1Cliente,
                                 Cliente = SA10.A1Nome,
                                 NF = SE10.E1Num,
                                 Valor = SE10.E1Valor,
                                 ValorSaldo = SE10.E1Saldo,
                                 DataEmissao = $"{SE10.E1Emissao.Substring(6, 2)}/{SE10.E1Emissao.Substring(4, 2)}/{SE10.E1Emissao.Substring(0, 4)}",
                                 DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}"
                             });

                Cliente = NReduz;
                if (NReduz != null)
                {
                    query = query.Where(x => x.Cliente == NReduz);
                }

                Relatorio = query.OrderByDescending(x => x.Valor).ToList();

                Clientes = (from SE10 in Protheus.Se1010s
                            join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                            where SE10.DELET != "*"
                            && (int)(object)SE10.E1Vencrea < (int)(object)DataInicio
                            && SA10.A1Clinter != "G" && (SE10.E1Baixa == "" || SE10.E1Saldo > 0)
                            && SE10.E1Tipo != "RA"
                            select new TitulosVencidos
                            {
                                CodCliente = SE10.E1Cliente,
                                Cliente = SA10.A1Nome,
                                Valor = SE10.E1Valor,
                                DataEmissao = $"{SE10.E1Emissao.Substring(6, 2)}/{SE10.E1Emissao.Substring(4, 2)}/{SE10.E1Emissao.Substring(0, 4)}",
                                DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}"
                            }).OrderBy(x => x.Cliente).Select(x => x.Cliente).Distinct().ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "TitulosVencidos Filter", user);
            }

            return Page();
        }

        public IActionResult OnPostExport(string NReduz)
        {
            try
            {

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                DateTime data = DateTime.Now;
                string DataInicio = $"{data.Year}{data.Month.ToString("D2")}{data.Day.ToString("D2")}";

                var query = (from SE10 in Protheus.Se1010s
                             join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                             where SE10.DELET != "*"
                             && (int)(object)SE10.E1Vencrea < (int)(object)DataInicio
                             && SA10.A1Clinter != "G" && (SE10.E1Baixa == "" || SE10.E1Saldo > 0)
                             && SE10.E1Tipo != "RA"
                             select new TitulosVencidos
                             {
                                 CodCliente = SE10.E1Cliente,
                                 Cliente = SA10.A1Nome,
                                 NF = SE10.E1Num,
                                 Valor = SE10.E1Valor,
                                 ValorSaldo = SE10.E1Saldo,
                                 DataEmissao = $"{SE10.E1Emissao.Substring(6, 2)}/{SE10.E1Emissao.Substring(4, 2)}/{SE10.E1Emissao.Substring(0, 4)}",
                                 DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}"
                             });
                Cliente = NReduz;

                if (NReduz != null)
                {
                    query = query.Where(x => x.Cliente == NReduz);
                }

                Relatorio = query.OrderByDescending(x => x.Valor).ToList();

                Clientes = (from SE10 in Protheus.Se1010s
                            join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                            where SE10.DELET != "*"
                            && (int)(object)SE10.E1Vencrea < (int)(object)DataInicio
                            && SA10.A1Clinter != "G" && (SE10.E1Baixa == "" || SE10.E1Saldo > 0)
                            && SE10.E1Tipo != "RA"
                            select new TitulosVencidos
                            {
                                CodCliente = SE10.E1Cliente,
                                Cliente = SA10.A1Nome,
                                Valor = SE10.E1Valor,
                                DataEmissao = $"{SE10.E1Emissao.Substring(6, 2)}/{SE10.E1Emissao.Substring(4, 2)}/{SE10.E1Emissao.Substring(0, 4)}",
                                DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}"
                            }).OrderBy(x => x.Cliente).Select(x => x.Cliente).Distinct().ToList();



                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Titulos Vencidos");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Titulos Vencidos");

                sheet.Cells[1, 1].Value = "Cliente";
                sheet.Cells[1, 2].Value = "NF";
                sheet.Cells[1, 3].Value = "Data Emissão";
                sheet.Cells[1, 4].Value = "Data de Vencimento";
                sheet.Cells[1, 5].Value = "Valor Original";
                sheet.Cells[1, 6].Value = "Saldo Devedor";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Cliente;
                    sheet.Cells[i, 2].Value = Pedido.NF;
                    sheet.Cells[i, 3].Value = Pedido.DataEmissao;
                    sheet.Cells[i, 4].Value = Pedido.DataVencimento;
                    sheet.Cells[i, 5].Value = Pedido.Valor;
                    sheet.Cells[i, 6].Value = Pedido.ValorSaldo;
                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TitulosVencidosInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "TitulosVencidos Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
