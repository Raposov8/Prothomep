using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Diretoria;
using OfficeOpenXml;

namespace SGID.Pages.Relatorios.Diretoria
{
    public class TitulosVencidosInterModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        private TOTVSINTERContext Protheus { get; set; }

        public List<TitulosVencidos> Relatorio { get; set; } = new List<TitulosVencidos>();

        public List<string> Clientes { get; set; } = new List<string>();
        public string Cliente { get; set; } = "";
        public string Tempo { get; set; } = "";
        public string Revelia { get; set; } = "";

        public TitulosVencidosInterModel(ApplicationDbContext sgid,TOTVSINTERContext inter)
        {
            SGID = sgid;
            Protheus = inter;
        }

        public void OnGet()
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                DateTime data = DateTime.Now;
                string DataInicio = $"{data.Year}{data.Month.ToString("D2")}{data.Day.ToString("D2")}";

                this.Tempo = "1";

                var query = (from SE10 in Protheus.Se1010s
                             join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                             join SC50 in Protheus.Sc5010s on SE10.E1Num equals SC50.C5Nota
                             where SE10.DELET != "*" && SA10.DELET!= "*" && SC50.DELET!="*"
                             && (int)(object)SE10.E1Vencrea > 20221231
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
                                 DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}",
                                 Revelia = SC50.C5Utpoper == "K" ? "REVELIA" : "PADRÃO"
                             });

                Relatorio = query.OrderByDescending(x => x.Cliente).ToList();

                Clientes = (from SE10 in Protheus.Se1010s
                            join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                            where SE10.DELET != "*" && SA10.DELET!= "*"
                            && (int)(object)SE10.E1Vencrea > 20221231
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
                Logger.Log(e, SGID, "TitulosVencidosInter", user);
            }
        }

        public IActionResult OnPost(string NReduz,string Tempo,string Revelia)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                DateTime data = DateTime.Now;
                string DataInicio = $"{data.Year}{data.Month.ToString("D2")}{data.Day.ToString("D2")}";


                if (Tempo == "1")
                {
                    var query = (from SE10 in Protheus.Se1010s
                                 join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                 join SC50 in Protheus.Sc5010s on SE10.E1Num equals SC50.C5Nota
                                 where SE10.DELET != "*" && SA10.DELET!= "*" && SC50.DELET!="*"
                                 && (int)(object)SE10.E1Vencrea > 20221231
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
                                     DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}",
                                     Revelia = SC50.C5Utpoper == "K" ? "REVELIA" : "PADRÃO"
                                 });

                    Cliente = NReduz;
                    if (NReduz != null)
                    {
                        query = query.Where(x => x.Cliente == NReduz);
                    }

                    Relatorio = query.OrderByDescending(x => x.Cliente).ToList();

                    Clientes = (from SE10 in Protheus.Se1010s
                                join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                where SE10.DELET != "*" && SA10.DELET!= "*" 
                                && (int)(object)SE10.E1Vencrea > 20221231
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
                else if (Tempo == "2")
                {
                    var query = (from SE10 in Protheus.Se1010s
                                 join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                 join SC50 in Protheus.Sc5010s on SE10.E1Num equals SC50.C5Nota
                                 where SE10.DELET != "*" && SA10.DELET!= "*" && SC50.DELET!="*"
                                 && (int)(object)SE10.E1Vencrea <= 20221231
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
                                     DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}",
                                     Revelia = SC50.C5Utpoper == "K" ? "REVELIA" : "PADRÃO"
                                 });

                    Cliente = NReduz;
                    if (NReduz != null)
                    {
                        query = query.Where(x => x.Cliente == NReduz);
                    }

                    Relatorio = query.OrderByDescending(x => x.Cliente).ToList();

                    Clientes = (from SE10 in Protheus.Se1010s
                                join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                where SE10.DELET != "*" && SA10.DELET!= "*" 
                                && (int)(object)SE10.E1Vencrea <= 20221231
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
                else
                {
                    var query = (from SE10 in Protheus.Se1010s
                                 join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                 join SC50 in Protheus.Sc5010s on SE10.E1Num equals SC50.C5Nota
                                 where SE10.DELET != "*" && SA10.DELET!= "*" && SC50.DELET!="*"
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
                                     DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}",
                                     Revelia = SC50.C5Utpoper == "K" ? "REVELIA" : "PADRÃO"
                                 });

                    Cliente = NReduz;
                    if (NReduz != null)
                    {
                        query = query.Where(x => x.Cliente == NReduz);
                    }

                    Relatorio = query.OrderByDescending(x => x.Cliente).ToList();

                    Clientes = (from SE10 in Protheus.Se1010s
                                join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                where SE10.DELET != "*" && SA10.DELET!= "*" 
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

                if (Revelia != "" && Revelia != null)
                {
                    Relatorio = Relatorio.Where(x => x.Revelia == Revelia).ToList();
                }

            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "TitulosVencidosInter Filter", user);
            }

            return Page();
        }

        public IActionResult OnPostExport(string NReduz,string Tempo,string Revelia)
        {
            try
            {

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                DateTime data = DateTime.Now;
                string DataInicio = $"{data.Year}{data.Month.ToString("D2")}{data.Day.ToString("D2")}";

                this.Tempo = Tempo;
                if (Tempo == "1")
                {
                    var query = (from SE10 in Protheus.Se1010s
                                 join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                 join SC50 in Protheus.Sc5010s on SE10.E1Num equals SC50.C5Nota
                                 where SE10.DELET != "*" && SA10.DELET!= "*" && SC50.DELET!="*"
                                 && (int)(object)SE10.E1Vencrea > 20221231
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
                                     DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}",
                                     Revelia = SC50.C5Utpoper == "K" ? "REVELIA" : "PADRÃO"
                                 });

                    Cliente = NReduz;
                    if (NReduz != null)
                    {
                        query = query.Where(x => x.Cliente == NReduz);
                    }

                    Relatorio = query.OrderByDescending(x => x.Cliente).ToList();

                    Clientes = (from SE10 in Protheus.Se1010s
                                join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                where SE10.DELET != "*" && SA10.DELET!= "*" 
                                && (int)(object)SE10.E1Vencrea > 20221231
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
                else if (Tempo == "2")
                {
                    var query = (from SE10 in Protheus.Se1010s
                                 join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                 join SC50 in Protheus.Sc5010s on SE10.E1Num equals SC50.C5Nota
                                 where SE10.DELET != "*" && SA10.DELET!= "*" && SC50.DELET!="*"
                                 && (int)(object)SE10.E1Vencrea <= 20221231
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
                                     DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}",
                                     Revelia = SC50.C5Utpoper == "K" ? "REVELIA" : "PADRÃO"
                                 });

                    Cliente = NReduz;
                    if (NReduz != null)
                    {
                        query = query.Where(x => x.Cliente == NReduz);
                    }

                    Relatorio = query.OrderByDescending(x => x.Cliente).ToList();

                    Clientes = (from SE10 in Protheus.Se1010s
                                join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                where SE10.DELET != "*" && SA10.DELET!= "*" 
                                && (int)(object)SE10.E1Vencrea <= 20221231
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
                else
                {
                    var query = (from SE10 in Protheus.Se1010s
                                 join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                 join SC50 in Protheus.Sc5010s on SE10.E1Num equals SC50.C5Nota
                                 where SE10.DELET != "*" && SA10.DELET!= "*" && SC50.DELET!="*"
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
                                     DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}",
                                     Revelia = SC50.C5Utpoper == "K" ? "REVELIA" : "PADRÃO"
                                 });

                    Cliente = NReduz;
                    if (NReduz != null)
                    {
                        query = query.Where(x => x.Cliente == NReduz);
                    }

                    Relatorio = query.OrderByDescending(x => x.Cliente).ToList();

                    Clientes = (from SE10 in Protheus.Se1010s
                                join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                                where SE10.DELET != "*" && SA10.DELET!= "*" 
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

                this.Revelia = Revelia;
                if (Revelia != "" && Revelia != null)
                {
                    Relatorio = Relatorio.Where(x => x.Revelia == Revelia).ToList();
                }

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
                Logger.Log(e, SGID, "TitulosVencidosInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
