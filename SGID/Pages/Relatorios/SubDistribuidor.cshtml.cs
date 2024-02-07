using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Denuo;

namespace SGID.Pages.Relatorios
{
    [Authorize]
    public class SubDistribuidorModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioSubDistribuidor> Relatorio { get; set; } = new List<RelatorioSubDistribuidor>();
        public double Total { get; set; }
        public double Desc { get; set; }
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }


        public SubDistribuidorModel(TOTVSDENUOContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }

        public void OnGet()
        {

        }

        public IActionResult OnPostAsync(DateTime Datainicio,DateTime DataFim, string NReduz)
        {
            try
            {
                Inicio = Datainicio;
                Fim = DataFim;

                var lista = (from SD20 in Protheus.Sd2010s
                             join SF40 in Protheus.Sf4010s on SD20.D2Tes equals SF40.F4Codigo
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, NumP = SD20.D2Pedido, Item = SD20.D2Itempv } equals new { Filial = SC60.C6Filial, NumP = SC60.C6Num, Item = SC60.C6Item }
                             join SC50 in Protheus.Sc5010s on new { Filial2 = SC60.C6Filial, NumC = SC60.C6Num } equals new { Filial2 = SC50.C5Filial, NumC = SC50.C5Num }
                             join SA10 in Protheus.Sa1010s on new { Codigo = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SD20.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*"
                             && SF40.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*"
                             && SF40.F4Duplic == "S" && SA10.A1Clinter == "S"
                             && (int)(object)SD20.D2Emissao >= (int)(object)Datainicio.ToString("yyyy/MM/dd").Replace("/", "") 
                             && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             orderby SD20.D2Emissao
                             select new RelatorioSubDistribuidor
                             {
                                 Nome = SA30.A3Nome,
                                 Nreduz = SA10.A1Nreduz,
                                 Doc = SD20.D2Doc,
                                 Cod = SB10.B1Cod,
                                 Desc = SB10.B1Desc,
                                 Quant = SD20.D2Quant,
                                 Total = SD20.D2Total,
                                 Descon = SD20.D2Descon,
                                 Emissao = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                                 Num = SC50.C5Num,
                                 Utpoper = SC50.C5Utpoper,
                                 Fabricante = SB10.B1Fabric
                             }
                            ).ToList();

                if (!string.IsNullOrEmpty(NReduz) && !string.IsNullOrWhiteSpace(NReduz))
                {
                    lista = lista.Where(x => x.Nreduz == NReduz).ToList();
                }

                Relatorio = lista;

                lista.ForEach(x => Total += x.Total);
                lista.ForEach(x => Desc += x.Descon);

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SubDistribuidor",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime Datainicio, DateTime DataFim, string NReduz)
        {
            try
            {
                var lista = (from SD20 in Protheus.Sd2010s
                             join SF40 in Protheus.Sf4010s on SD20.D2Tes equals SF40.F4Codigo
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, NumP = SD20.D2Pedido, Item = SD20.D2Itempv } equals new { Filial = SC60.C6Filial, NumP = SC60.C6Num, Item = SC60.C6Item }
                             join SC50 in Protheus.Sc5010s on new { Filial2 = SC60.C6Filial, NumC = SC60.C6Num } equals new { Filial2 = SC50.C5Filial, NumC = SC50.C5Num }
                             join SA10 in Protheus.Sa1010s on new { Codigo = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SD20.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*"
                             && SF40.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*"
                             && SF40.F4Duplic == "S" && SA10.A1Clinter == "S"
                             && Convert.ToInt32(SD20.D2Emissao) >= Convert.ToInt32(Datainicio.ToString("yyyy/MM/dd").Replace("/", "")) 
                             && Convert.ToInt32(SD20.D2Emissao) <= Convert.ToInt32(DataFim.ToString("yyyy/MM/dd").Replace("/", ""))
                             orderby SD20.D2Emissao
                             select new RelatorioSubDistribuidor
                             {
                                 Nome = SA30.A3Nome,
                                 Nreduz = SA10.A1Nreduz,
                                 Doc = SD20.D2Doc,
                                 Cod = SB10.B1Cod,
                                 Desc = SB10.B1Desc,
                                 Quant = SD20.D2Quant,
                                 Total = SD20.D2Total,
                                 Descon = SD20.D2Descon,
                                 Emissao = SD20.D2Emissao,
                                 Num = SC50.C5Num,
                                 Utpoper = SC50.C5Utpoper,
                                 Fabricante = SB10.B1Fabric
                             }
                            ).ToList();

                if (!string.IsNullOrEmpty(NReduz) && !string.IsNullOrWhiteSpace(NReduz))
                {
                    lista = lista.Where(x => x.Nreduz == NReduz).ToList();
                }

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("SubDistribuidor Faturado");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "SubDistribuidor Faturado");

                sheet.Cells[1, 1].Value = "Nome";
                sheet.Cells[1, 2].Value = "Nome Cliente";
                sheet.Cells[1, 3].Value = "DOC";
                sheet.Cells[1, 4].Value = "COD";
                sheet.Cells[1, 5].Value = "Produto";
                sheet.Cells[1, 6].Value = "QTD";
                sheet.Cells[1, 7].Value = "Valor";
                sheet.Cells[1, 8].Value = "Desconto";
                sheet.Cells[1, 9].Value = "Emissao";
                sheet.Cells[1, 10].Value = "Num";
                sheet.Cells[1, 11].Value = "UTPOPER";
                sheet.Cells[1, 12].Value = "Fabricante";

                int i = 2;

                lista.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Nome;
                    sheet.Cells[i, 2].Value = Pedido.Nreduz;
                    sheet.Cells[i, 3].Value = Pedido.Doc;
                    sheet.Cells[i, 4].Value = Pedido.Cod;
                    sheet.Cells[i, 5].Value = Pedido.Desc;
                    sheet.Cells[i, 6].Value = Pedido.Quant;
                    sheet.Cells[i, 7].Value = Pedido.Total;
                    sheet.Cells[i, 8].Value = Pedido.Descon;
                    sheet.Cells[i, 9].Value = $"{Pedido.Emissao.Substring(6, 2)}/{Pedido.Emissao.Substring(4, 2)}/{Pedido.Emissao.Substring(0, 4)}";
                    sheet.Cells[i, 10].Value = Pedido.Num;
                    sheet.Cells[i, 11].Value = Pedido.Utpoper;
                    sheet.Cells[i, 12].Value = Pedido.Fabricante;

                    Total += Pedido.Total;
                    Desc += Pedido.Descon;

                    i++;
                });

                sheet.Cells[i, 1].Value = "Total Geral";
                sheet.Cells[i, 7].Value = Total.ToString("0.##");
                sheet.Cells[i, 8].Value = Desc.ToString("0.##");

                sheet.Cells[i, 1, i, 6].Merge = true;

                var format = sheet.Cells[i, 3, i, 4];
                format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SubDistribuidorFaturadoDenuo.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SubDistribuidor Excel",user);

                return LocalRedirect("/error");
            }
        }
    }
}
