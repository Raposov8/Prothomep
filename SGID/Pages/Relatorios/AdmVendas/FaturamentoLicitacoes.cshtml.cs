using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Denuo;

namespace SGID.Pages.Relatorios.AdmVendas
{
    [Authorize]
    public class FaturamentoLicitacoesModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioCirurgiasFaturadas> Relatorio = new List<RelatorioCirurgiasFaturadas>();

        public FaturamentoLicitacoesModel(TOTVSDENUOContext context, ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                int[] CF = new int[] { 5551, 6551, 6107, 6109 };
                var user = User.Identity.Name.Split("@")[0].ToUpper();

                var query = (from SD20 in Protheus.Sd2010s
                             join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Codigo = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Codigo = SC60.C6Produto }
                             where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                             && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                             ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", ""))
                             && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SC50.C5Xtipopv != "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
                             && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
                             select new
                             {
                                 Filial = SD20.D2Filial,
                                 Clifor = SD20.D2Cliente,
                                 Loja = SD20.D2Loja,
                                 Nome = SA10.A1Nome,
                                 Tipo = SA10.A1Clinter,
                                 NF = SD20.D2Doc,
                                 Serie = SD20.D2Serie,
                                 Emissao = SD20.D2Emissao,
                                 Pedido = SD20.D2Pedido,
                                 Total = SD20.D2Total,
                                 Valipi = SD20.D2Valipi,
                                 Valicm = SD20.D2Valicm,
                                 Descon = SD20.D2Descon,
                                 Unumage = SC50.C5Unumage,
                                 SC50.C5Emissao,
                                 SA30.A3Nome,
                                 DataCirurgia = SC50.C5XDtcir,
                                 NomMed = SC50.C5XNmmed,
                                 NomPac = SC50.C5XNmpac,
                                 NomPla = SC50.C5XNmpla,
                                 SC50.C5Utpoper,
                                 SD20.D2Cod,
                                 SB10.B1Desc,
                                 SC60.C6Xitlici
                             });


                Relatorio = query.GroupBy(x => new
                {
                    x.Filial,
                    x.Clifor,
                    x.Loja,
                    x.Nome,
                    x.Tipo,
                    x.NF,
                    x.Serie,
                    x.Emissao,
                    x.Pedido,
                    x.Unumage,
                    x.C5Emissao,
                    x.A3Nome,
                    x.DataCirurgia,
                    x.NomMed,
                    x.NomPac,
                    x.NomPla,
                    x.C5Utpoper,
                    x.D2Cod,
                    x.B1Desc,
                    x.C6Xitlici
                }).Select(x => new RelatorioCirurgiasFaturadas
                {
                    Filial = x.Key.Filial,
                    Clifor = x.Key.Clifor,
                    Loja = x.Key.Loja,
                    Nome = x.Key.Nome,
                    Tipo = x.Key.Tipo,
                    Nf = x.Key.NF,
                    Serie = x.Key.Serie,
                    Emissao = $"{x.Key.Emissao.Substring(6, 2)}/{x.Key.Emissao.Substring(4, 2)}/{x.Key.Emissao.Substring(0, 4)}",
                    Pedido = x.Key.Pedido,
                    Total = x.Sum(c => c.Total),
                    Valipi = x.Sum(c => c.Valipi),
                    Valicm = x.Sum(c => c.Valicm),
                    Descon = x.Sum(c => c.Descon),
                    Unumage = x.Key.Unumage,
                    C5Emissao = x.Key.C5Emissao,
                    A3Nome = x.Key.A3Nome,
                    XDtcir = $"{x.Key.DataCirurgia.Substring(6, 2)}/{x.Key.DataCirurgia.Substring(4, 2)}/{x.Key.DataCirurgia.Substring(0, 4)}",
                    XNMMed = x.Key.NomMed,
                    XNMPac = x.Key.NomPac,
                    XNMPla = x.Key.NomPla,
                    Utpoper = x.Key.C5Utpoper,
                    D2Cod = x.Key.D2Cod,
                    B1Desc = x.Key.B1Desc,
                    LicitacaoCodigo = x.Key.C6Xitlici
                }).OrderBy(x => x.A3Nome).ToList();



                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoLicitacoesADM", user);
                return LocalRedirect("/error");
            }
        }
        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Relatorio = new List<RelatorioCirurgiasFaturadas>();
                int[] CF = new int[] { 5551, 6551, 6107, 6109 };
                var user = User.Identity.Name.Split("@")[0].ToUpper();

                var query = (from SD20 in Protheus.Sd2010s
                             join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial,Num=SD20.D2Pedido,Codigo=SD20.D2Cod} equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Codigo = SC60.C6Produto }
                             where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                             && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                             ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", ""))
                             && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SC50.C5Xtipopv != "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
                             && (SA10.A1Xgrinte == "000011" || SA10.A1Xgrinte == "000012")
                             select new
                             {
                                 Filial = SD20.D2Filial,
                                 Clifor = SD20.D2Cliente,
                                 Loja = SD20.D2Loja,
                                 Nome = SA10.A1Nome,
                                 Tipo = SA10.A1Clinter,
                                 NF = SD20.D2Doc,
                                 Serie = SD20.D2Serie,
                                 Emissao = SD20.D2Emissao,
                                 Pedido = SD20.D2Pedido,
                                 Total = SD20.D2Total,
                                 Valipi = SD20.D2Valipi,
                                 Valicm = SD20.D2Valicm,
                                 Descon = SD20.D2Descon,
                                 Unumage = SC50.C5Unumage,
                                 SC50.C5Emissao,
                                 SA30.A3Nome,
                                 DataCirurgia = SC50.C5XDtcir,
                                 NomMed = SC50.C5XNmmed,
                                 NomPac = SC50.C5XNmpac,
                                 NomPla = SC50.C5XNmpla,
                                 SC50.C5Utpoper,
                                 SD20.D2Cod,
                                 SB10.B1Desc,
                                 SC60.C6Xitlici
                             });


                Relatorio = query.GroupBy(x => new
                {
                    x.Filial,
                    x.Clifor,
                    x.Loja,
                    x.Nome,
                    x.Tipo,
                    x.NF,
                    x.Serie,
                    x.Emissao,
                    x.Pedido,
                    x.Unumage,
                    x.C5Emissao,
                    x.A3Nome,
                    x.DataCirurgia,
                    x.NomMed,
                    x.NomPac,
                    x.NomPla,
                    x.C5Utpoper,
                    x.D2Cod,
                    x.B1Desc,
                    x.C6Xitlici
                }).Select(x => new RelatorioCirurgiasFaturadas
                {
                    Filial = x.Key.Filial,
                    Clifor = x.Key.Clifor,
                    Loja = x.Key.Loja,
                    Nome = x.Key.Nome,
                    Tipo = x.Key.Tipo,
                    Nf = x.Key.NF,
                    Serie = x.Key.Serie,
                    Emissao = $"{x.Key.Emissao.Substring(6, 2)}/{x.Key.Emissao.Substring(4, 2)}/{x.Key.Emissao.Substring(0, 4)}",
                    Pedido = x.Key.Pedido,
                    Total = x.Sum(c => c.Total),
                    Valipi = x.Sum(c => c.Valipi),
                    Valicm = x.Sum(c => c.Valicm),
                    Descon = x.Sum(c => c.Descon),
                    Unumage = x.Key.Unumage,
                    C5Emissao = x.Key.C5Emissao,
                    A3Nome = x.Key.A3Nome,
                    XDtcir = $"{x.Key.DataCirurgia.Substring(6, 2)}/{x.Key.DataCirurgia.Substring(4, 2)}/{x.Key.DataCirurgia.Substring(0, 4)}",
                    XNMMed = x.Key.NomMed,
                    XNMPac = x.Key.NomPac,
                    XNMPla = x.Key.NomPla,
                    Utpoper = x.Key.C5Utpoper,
                    D2Cod = x.Key.D2Cod,
                    B1Desc = x.Key.B1Desc,
                    LicitacaoCodigo = x.Key.C6Xitlici
                }).OrderBy(x => x.A3Nome).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Cirurgias Faturadas");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Cirurgias Faturadas");

                sheet.Cells[1, 1].Value = "Pedido";
                sheet.Cells[1, 2].Value = "Data Cirurgia";
                sheet.Cells[1, 3].Value = "NF";
                sheet.Cells[1, 4].Value = "Emissão NF";
                sheet.Cells[1, 5].Value = "Nome";
                sheet.Cells[1, 6].Value = "Vendedor";
                sheet.Cells[1, 7].Value = "Médico";
                sheet.Cells[1, 8].Value = "Paciente";
                sheet.Cells[1, 9].Value = "Convênio";
                sheet.Cells[1, 10].Value = "Produto";
                sheet.Cells[1, 11].Value = "Produto Licitacao";
                sheet.Cells[1, 12].Value = "Desc.Produto";
                sheet.Cells[1, 13].Value = "Total";     

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Pedido;
                    sheet.Cells[i, 2].Value = Pedido.XDtcir;
                    sheet.Cells[i, 3].Value = Pedido.Nf;
                    sheet.Cells[i, 4].Value = Pedido.Emissao;
                    sheet.Cells[i, 5].Value = Pedido.Nome;
                    sheet.Cells[i, 6].Value = Pedido.A3Nome;
                    sheet.Cells[i, 7].Value = Pedido.XNMMed;
                    sheet.Cells[i, 8].Value = Pedido.XNMPac;
                    sheet.Cells[i, 9].Value = Pedido.XNMPla;
                    sheet.Cells[i, 10].Value = Pedido.D2Cod;
                    sheet.Cells[i, 11].Value = Pedido.LicitacaoCodigo;
                    sheet.Cells[i, 12].Value = Pedido.B1Desc;
                    sheet.Cells[i, 13].Value = Pedido.Total;

                    i++;
                });

                var format = sheet.Cells[i, 8, i, 9];
                format.Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FaturamentoLicitacoes.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoLicitacoesADM Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
