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

                var query = (from SD20 in Protheus.Sd2010s
                             join SF40 in Protheus.Sf4010s on new { Filial = SD20.D2Filial, Cod = SD20.D2Tes } equals new { Filial = SF40.F4Filial, Cod = SF40.F4Codigo }
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Prod = SD20.D2Cod, Item = SD20.D2Itempv } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Prod = SC60.C6Produto, Item = SC60.C6Item }
                             join SC50 in Protheus.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             where SD20.DELET != "*" && SF40.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*" && SB10.DELET != "*"
                             && SC50.C5Utpoper != "" && SD20.D2Quant > SD20.D2Qtdedev && SC50.C5Utpoper != "F" && SB10.B1Tipo != "AI"
                             && (int)(object)SD20.D2Dtvalid >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Dtvalid <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             select new RelatorioValidadePermanente
                             {
                                 Filial = SD20.D2Filial,
                                 Codigo = SD20.D2Cod,
                                 Lote = SD20.D2Lotectl,
                                 Processo = SC50.C5Uproces,
                                 Agendamento = SC50.C5Unumage,
                                 Patrimonio = SC60.C6Upatrim,
                                 ValidLote = $"{SD20.D2Dtvalid.Substring(6, 2)}/{SD20.D2Dtvalid.Substring(4, 2)}/{SD20.D2Dtvalid.Substring(0, 4)}",
                                 Operacao = SC50.C5Utpoper == "D" ? "DIARIA" : SC50.C5Utpoper == "P" ? "PERMANENTE" : SC50.C5Utpoper == "C" ? "CONGRESSO" : SC50.C5Utpoper == "E" ? "EMPRESTIMO" : SC50.C5Utpoper == "B" ? "CONSIGNACAO" : SC50.C5Utpoper == "X" ? "DEMONSTRACAO" : SC50.C5Utpoper == "F" ? "FATURAMENTO" : "OUTROS",
                                 NFSaida = SD20.D2Doc,
                                 SerieSaida = SD20.D2Serie,
                                 EmissaoNf = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                                 ItSaida = SD20.D2Item,
                                 QtdSaida = SD20.D2Quant,
                                 Saldo = 0.00
                             }).ToList();

                var query2 = (
                           from SD20 in Protheus.Sd2010s
                           join SB60 in Protheus.Sb6010s on new { Filial = SD20.D2Filial, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Cod = SD20.D2Cod, Ident = SD20.D2Identb6 } equals new { Filial = SB60.B6Filial, Cli = SB60.B6Clifor, Loja = SB60.B6Loja, Cod = SB60.B6Produto, Ident = SB60.B6Ident }
                           join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Prod = SD20.D2Cod, Item = SD20.D2Itempv } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Prod = SC60.C6Produto, Item = SC60.C6Item }
                           join SC50 in Protheus.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                           join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                           where SD20.DELET != "*" && SB60.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*" && SB10.DELET != "*"
                           && SB10.B1Tipo != "AI" && SB60.B6Saldo != 0
                           && (int)(object)SD20.D2Dtvalid >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Dtvalid <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                           select new RelatorioValidadePermanente
                           {
                               Filial = SD20.D2Filial,
                               Codigo = SD20.D2Cod,
                               Lote = SD20.D2Lotectl,
                               Processo = SC50.C5Uproces,
                               Agendamento = SC50.C5Unumage,
                               Patrimonio = SC60.C6Upatrim,
                               ValidLote = $"{SD20.D2Dtvalid.Substring(6, 2)}/{SD20.D2Dtvalid.Substring(4, 2)}/{SD20.D2Dtvalid.Substring(0, 4)}",
                               Operacao = SC50.C5Utpoper == "D" ? "DIARIA" : SC50.C5Utpoper == "P" ? "PERMANENTE" : SC50.C5Utpoper == "C" ? "CONGRESSO" : SC50.C5Utpoper == "E" ? "EMPRESTIMO" : SC50.C5Utpoper == "B" ? "CONSIGNACAO" : SC50.C5Utpoper == "X" ? "DEMONSTRACAO" : SC50.C5Utpoper == "F" ? "FATURAMENTO" : "OUTROS",
                               NFSaida = SD20.D2Doc,
                               SerieSaida = SD20.D2Serie,
                               EmissaoNf = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                               ItSaida = SD20.D2Item,
                               QtdSaida = SD20.D2Quant,
                               Saldo = SB60.B6Saldo
                           }).ToList();

                Relatorios = query.Concat(query2).OrderBy(x => x.ValidLote).ToList();

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

                var query = (from SD20 in Protheus.Sd2010s
                             join SF40 in Protheus.Sf4010s on new { Filial = SD20.D2Filial, Cod = SD20.D2Tes } equals new { Filial = SF40.F4Filial, Cod = SF40.F4Codigo }
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Prod = SD20.D2Cod, Item = SD20.D2Itempv } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Prod = SC60.C6Produto, Item = SC60.C6Item }
                             join SC50 in Protheus.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             where SD20.DELET != "*" && SF40.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*" && SB10.DELET != "*"
                             && SC50.C5Utpoper != "" && SD20.D2Quant > SD20.D2Qtdedev && SC50.C5Utpoper != "F" && SB10.B1Tipo != "AI"
                             && (int)(object)SD20.D2Dtvalid >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Dtvalid <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             select new RelatorioValidadePermanente
                             {
                                 Filial = SD20.D2Filial,
                                 Codigo = SD20.D2Cod,
                                 Lote = SD20.D2Lotectl,
                                 Processo = SC50.C5Uproces,
                                 Agendamento = SC50.C5Unumage,
                                 Patrimonio = SC60.C6Upatrim,
                                 ValidLote = $"{SD20.D2Dtvalid.Substring(6, 2)}/{SD20.D2Dtvalid.Substring(4, 2)}/{SD20.D2Dtvalid.Substring(0, 4)}",
                                 Operacao = SC50.C5Utpoper == "D" ? "DIARIA" : SC50.C5Utpoper == "P" ? "PERMANENTE" : SC50.C5Utpoper == "C" ? "CONGRESSO" : SC50.C5Utpoper == "E" ? "EMPRESTIMO" : SC50.C5Utpoper == "B" ? "CONSIGNACAO" : SC50.C5Utpoper == "X" ? "DEMONSTRACAO" : SC50.C5Utpoper == "F" ? "FATURAMENTO" : "OUTROS",
                                 NFSaida = SD20.D2Doc,
                                 SerieSaida = SD20.D2Serie,
                                 EmissaoNf = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                                 ItSaida = SD20.D2Item,
                                 QtdSaida = SD20.D2Quant,
                                 Saldo = 0.00
                             }
                             ).ToList();
                var query2 = (
                           from SD20 in Protheus.Sd2010s
                           join SB60 in Protheus.Sb6010s on new { Filial = SD20.D2Filial, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Cod = SD20.D2Cod, Ident = SD20.D2Identb6 } equals new { Filial = SB60.B6Filial, Cli = SB60.B6Clifor, Loja = SB60.B6Loja, Cod = SB60.B6Produto, Ident = SB60.B6Ident }
                           join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Prod = SD20.D2Cod, Item = SD20.D2Itempv } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Prod = SC60.C6Produto, Item = SC60.C6Item }
                           join SC50 in Protheus.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                           join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                           where SD20.DELET != "*" && SB60.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*" && SB10.DELET != "*"
                           && SB10.B1Tipo != "AI" && SB60.B6Saldo != 0
                           && (int)(object)SD20.D2Dtvalid >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Dtvalid <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                           select new RelatorioValidadePermanente
                           {
                               Filial = SD20.D2Filial,
                               Codigo = SD20.D2Cod,
                               Lote = SD20.D2Lotectl,
                               Processo = SC50.C5Uproces,
                               Agendamento = SC50.C5Unumage,
                               Patrimonio = SC60.C6Upatrim,
                               ValidLote = $"{SD20.D2Dtvalid.Substring(6, 2)}/{SD20.D2Dtvalid.Substring(4, 2)}/{SD20.D2Dtvalid.Substring(0, 4)}",
                               Operacao = SC50.C5Utpoper == "D" ? "DIARIA" : SC50.C5Utpoper == "P" ? "PERMANENTE" : SC50.C5Utpoper == "C" ? "CONGRESSO" : SC50.C5Utpoper == "E" ? "EMPRESTIMO" : SC50.C5Utpoper == "B" ? "CONSIGNACAO" : SC50.C5Utpoper == "X" ? "DEMONSTRACAO" : SC50.C5Utpoper == "F" ? "FATURAMENTO" : "OUTROS",
                               NFSaida = SD20.D2Doc,
                               SerieSaida = SD20.D2Serie,
                               EmissaoNf = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                               ItSaida = SD20.D2Item,
                               QtdSaida = SD20.D2Quant,
                               Saldo = SB60.B6Saldo
                           }
                           ).ToList();

                Relatorios = query.Concat(query2).OrderBy(x => x.ValidLote).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("ValidadePermanente");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "ValidadePermanente");

                sheet.Cells[1, 1].Value = "Filial";
                sheet.Cells[1, 2].Value = "Codigo";
                sheet.Cells[1, 3].Value = "Lote";
                sheet.Cells[1, 4].Value = "Processo";
                sheet.Cells[1, 5].Value = "Agendamento";
                sheet.Cells[1, 6].Value = "Patrimonio";
                sheet.Cells[1, 7].Value = "Valid Lote";
                sheet.Cells[1, 8].Value = "Operacao";
                sheet.Cells[1, 9].Value = "NF Saida";
                sheet.Cells[1, 10].Value = "Serie Saida";
                sheet.Cells[1, 11].Value = "Emissao NF";
                sheet.Cells[1, 12].Value = "IT Saida";
                sheet.Cells[1, 13].Value = "QTD Saida";
                sheet.Cells[1, 14].Value = "Saldo";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.Codigo;
                    sheet.Cells[i, 3].Value = Pedido.Lote;
                    sheet.Cells[i, 4].Value = Pedido.Processo;
                    sheet.Cells[i, 5].Value = Pedido.Agendamento;
                    sheet.Cells[i, 6].Value = Pedido.Patrimonio;
                    sheet.Cells[i, 7].Value = Pedido.ValidLote;
                    sheet.Cells[i, 8].Value = Pedido.Operacao;
                    sheet.Cells[i, 9].Value = Pedido.NFSaida;
                    sheet.Cells[i, 10].Value = Pedido.SerieSaida;
                    sheet.Cells[i, 11].Value = Pedido.EmissaoNf;
                    sheet.Cells[i, 12].Value = Pedido.ItSaida;
                    sheet.Cells[i, 13].Value = Pedido.QtdSaida;
                    sheet.Cells[i, 14].Value = Pedido.Saldo;

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
