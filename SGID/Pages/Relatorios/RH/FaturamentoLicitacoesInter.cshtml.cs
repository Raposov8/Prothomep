using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.RH
{
    [Authorize]
    public class FaturamentoLicitacoesInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioCirurgiasFaturadas> Relatorio = new List<RelatorioCirurgiasFaturadas>();

        public FaturamentoLicitacoesInterModel(TOTVSINTERContext context, ApplicationDbContext sgid)
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

                string[] CF = new string[] { "5551", "6551", "6107", "6109" };

                var DataI = DataInicio.ToString("yyyy/MM/dd").Replace("/", "");
                var DataF = DataFim.ToString("yyyy/MM/dd").Replace("/", "");

                //

                var query = (from SD20 in Protheus.Sd2010s
                             join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Pedido = SC60.C6Num, Cliente = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                             where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SC60.DELET != "*"
                             && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                             ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains(SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataI && (int)(object)SD20.D2Emissao <= (int)(object)DataF)
                             && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
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
                                 SC50.C5Nomvend,
                                 DataCirurgia = SC50.C5XDtcir,
                                 NomMed = SC50.C5XNmmed,
                                 NomPac = SC50.C5XNmpac,
                                 NomPla = SC50.C5XNmpla,
                                 SC50.C5Utpoper,
                                 SD20.D2Cod,
                                 SB10.B1Desc,
                                 SC60.C6Produto
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
                    x.C5Nomvend,
                    x.DataCirurgia,
                    x.NomMed,
                    x.NomPac,
                    x.NomPla,
                    x.C5Utpoper,
                    x.D2Cod,
                    x.B1Desc,
                    x.C6Produto
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
                    A3Nome = x.Key.C5Nomvend,
                    XDtcir = $"{x.Key.DataCirurgia.Substring(6, 2)}/{x.Key.DataCirurgia.Substring(4, 2)}/{x.Key.DataCirurgia.Substring(0, 4)}",
                    XNMMed = x.Key.NomMed,
                    XNMPac = x.Key.NomPac,
                    XNMPla = x.Key.NomPla,
                    Utpoper = x.Key.C5Utpoper,
                    D2Cod = x.Key.D2Cod,
                    B1Desc = x.Key.B1Desc

                }).OrderBy(x => x.A3Nome).ToList();

                var soma = 0.00;

                Relatorio.ForEach(x => soma += x.Total);

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoLicitacoesInter", user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Relatorio = new List<RelatorioCirurgiasFaturadas>();
                string[] CF = new string[] { "5551", "6551", "6107", "6109" };
                var user = User.Identity.Name.Split("@")[0].ToUpper();

                var DataI = DataInicio.ToString("yyyy/MM/dd").Replace("/", "");
                var DataF = DataFim.ToString("yyyy/MM/dd").Replace("/", "");

                //&& (SA10.A1Xgrinte != "000011" || SA10.A1Xgrinte != "000012")

                var query = (from SD20 in Protheus.Sd2010s
                             join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Pedido = SD20.D2Pedido, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Pedido = SC60.C6Num, Cliente = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                             where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SC60.DELET != "*"
                             && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                             ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains(SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataI && (int)(object)SD20.D2Emissao <= (int)(object)DataF)
                             && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200701
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
                                 SC50.C5Nomvend,
                                 DataCirurgia = SC50.C5XDtcir,
                                 NomMed = SC50.C5XNmmed,
                                 NomPac = SC50.C5XNmpac,
                                 NomPla = SC50.C5XNmpla,
                                 SC50.C5Utpoper,
                                 SD20.D2Cod,
                                 SB10.B1Desc,
                                 SC60.C6Produto
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
                    x.C5Nomvend,
                    x.DataCirurgia,
                    x.NomMed,
                    x.NomPac,
                    x.NomPla,
                    x.C5Utpoper,
                    x.D2Cod,
                    x.B1Desc,
                    x.C6Produto
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
                    A3Nome = x.Key.C5Nomvend,
                    XDtcir = $"{x.Key.DataCirurgia.Substring(6, 2)}/{x.Key.DataCirurgia.Substring(4, 2)}/{x.Key.DataCirurgia.Substring(0, 4)}",
                    XNMMed = x.Key.NomMed,
                    XNMPac = x.Key.NomPac,
                    XNMPla = x.Key.NomPla,
                    Utpoper = x.Key.C5Utpoper,
                    D2Cod = x.Key.D2Cod,
                    B1Desc = x.Key.B1Desc

                }).OrderBy(x => x.A3Nome).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Cirurgias Faturadas Inter");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Cirurgias Faturadas Inter");

                sheet.Cells[1, 1].Value = "Filial";
                sheet.Cells[1, 2].Value = "Pedido";
                sheet.Cells[1, 3].Value = "Agendamento";
                sheet.Cells[1, 4].Value = "Data Cirurgia";
                sheet.Cells[1, 5].Value = "NF";
                sheet.Cells[1, 6].Value = "Serie";
                sheet.Cells[1, 7].Value = "Emissão NF";
                sheet.Cells[1, 8].Value = "Total";
                sheet.Cells[1, 9].Value = "Descon";
                sheet.Cells[1, 10].Value = "CliFor";
                sheet.Cells[1, 11].Value = "Loja";
                sheet.Cells[1, 12].Value = "Nome";
                sheet.Cells[1, 13].Value = "Vendedor";
                sheet.Cells[1, 14].Value = "Médico";
                sheet.Cells[1, 15].Value = "Paciente";
                sheet.Cells[1, 16].Value = "Convênio";
                sheet.Cells[1, 17].Value = "Produto";
                sheet.Cells[1, 18].Value = "Desc. Produto";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.Pedido;
                    sheet.Cells[i, 3].Value = Pedido.Unumage;
                    sheet.Cells[i, 4].Value = Pedido.XDtcir;
                    sheet.Cells[i, 5].Value = Pedido.Nf;
                    sheet.Cells[i, 6].Value = Pedido.Serie;
                    sheet.Cells[i, 7].Value = Pedido.Emissao;
                    sheet.Cells[i, 8].Value = Pedido.Total;
                    sheet.Cells[i, 9].Value = Pedido.Descon;
                    sheet.Cells[i, 10].Value = Pedido.Clifor;
                    sheet.Cells[i, 11].Value = Pedido.Loja;
                    sheet.Cells[i, 12].Value = Pedido.Nome;
                    sheet.Cells[i, 13].Value = Pedido.A3Nome;
                    sheet.Cells[i, 14].Value = Pedido.XNMMed;
                    sheet.Cells[i, 15].Value = Pedido.XNMPac;
                    sheet.Cells[i, 16].Value = Pedido.XNMPla;
                    sheet.Cells[i, 17].Value = Pedido.D2Cod;
                    sheet.Cells[i, 18].Value = Pedido.B1Desc;

                    i++;
                });

                var format = sheet.Cells[i, 8, i, 9];
                format.Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FaturamentoLicitacoesInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoLicitacoesInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
