using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Qualidade;

namespace SGID.Pages.Relatorios.Qualidade
{
    [Authorize]
    public class ValidadePermanenteInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioValidadePermanente> Relatorios { get; set; } = new List<RelatorioValidadePermanente>();

        public ValidadePermanenteInterModel(TOTVSINTERContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
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

                var query = (from SD20 in Protheus.Sd2010s
                             join SF40 in Protheus.Sf4010s on new { Filial = SD20.D2Filial, Cod = SD20.D2Tes } equals new { Filial = SF40.F4Filial, Cod = SF40.F4Codigo } into Sa
                             from a in Sa.DefaultIfEmpty()
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Prod = SD20.D2Cod, Item = SD20.D2Itempv } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Prod = SC60.C6Produto, Item = SC60.C6Item } into Se
                             from c in Se.DefaultIfEmpty()
                             join SC50 in Protheus.Sc5010s on new { Filial = c.C6Filial, Num = c.C6Num } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num } into Si
                             from b in Si.DefaultIfEmpty()
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod into So
                             from d in So.DefaultIfEmpty()
                             join SD10 in Protheus.Sd1010s on SD20.D2Doc equals SD10.D1Nfori into Su
                             from e in Su.DefaultIfEmpty()
                             where SD20.DELET != "*" && a.DELET != "*" && c.DELET != "*" && b.DELET != "*" && d.DELET != "*"
                             && b.C5Utpoper != "" && SD20.D2Quant > SD20.D2Qtdedev && b.C5Utpoper != "F" && d.B1Tipo != "AI"
                             && (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Emissao<= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             && e.D1Nfori == null
                             select new RelatorioValidadePermanente
                             {
                                 Filial = SD20.D2Filial,
                                 Produto = SD20.D2Cod,
                                 Lote = SD20.D2Lotectl,
                                 Agendamento = b.C5Unumage,
                                 Validade = $"{SD20.D2Dtvalid.Substring(6, 2)}/{SD20.D2Dtvalid.Substring(4, 2)}/{SD20.D2Dtvalid.Substring(0, 4)}",
                                 Operacoes = b.C5Utpoper == "D" ? "DIARIA" : b.C5Utpoper == "P" ? "PERMANENTE" : b.C5Utpoper == "C" ? "CONGRESSO" : b.C5Utpoper == "E" ? "EMPRESTIMO" : b.C5Utpoper == "B" ? "CONSIGNACAO" : b.C5Utpoper == "X" ? "DEMONSTRACAO" : b.C5Utpoper == "F" ? "FATURAMENTO" : "OUTROS",
                                 NF = SD20.D2Doc,
                                 EmissaoNF = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                                 QTD = SD20.D2Quant,
                                 Pedido = b.C5Num,
                                 Cliente = b.C5Nomcli,
                                 Patrimonio = c.C6Upatrim
                             }).ToList();



                Relatorios = query.OrderBy(x => x.EmissaoNF).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ValidadePermanente", user);

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
                             join SF40 in Protheus.Sf4010s on new { Filial = SD20.D2Filial, Cod = SD20.D2Tes } equals new { Filial = SF40.F4Filial, Cod = SF40.F4Codigo } into Sa
                             from a in Sa.DefaultIfEmpty()
                             join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Prod = SD20.D2Cod, Item = SD20.D2Itempv } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Prod = SC60.C6Produto, Item = SC60.C6Item } into Se
                             from c in Se.DefaultIfEmpty()
                             join SC50 in Protheus.Sc5010s on new { Filial = c.C6Filial, Num = c.C6Num } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num } into Si
                             from b in Si.DefaultIfEmpty()
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod into So
                             from d in So.DefaultIfEmpty()
                             join SD10 in Protheus.Sd1010s on SD20.D2Doc equals SD10.D1Nfori into Su
                             from e in Su.DefaultIfEmpty()
                             where SD20.DELET != "*" && a.DELET != "*" && c.DELET != "*" && b.DELET != "*" && d.DELET != "*"
                             && b.C5Utpoper != "" && SD20.D2Quant > SD20.D2Qtdedev && b.C5Utpoper != "F" && d.B1Tipo != "AI"
                             && (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             && e.D1Nfori == null
                             select new RelatorioValidadePermanente
                             {
                                 Filial = SD20.D2Filial,
                                 Produto = SD20.D2Cod,
                                 Lote = SD20.D2Lotectl,
                                 Agendamento = b.C5Unumage,
                                 Validade = $"{SD20.D2Dtvalid.Substring(6, 2)}/{SD20.D2Dtvalid.Substring(4, 2)}/{SD20.D2Dtvalid.Substring(0, 4)}",
                                 Operacoes = b.C5Utpoper == "D" ? "DIARIA" : b.C5Utpoper == "P" ? "PERMANENTE" : b.C5Utpoper == "C" ? "CONGRESSO" : b.C5Utpoper == "E" ? "EMPRESTIMO" : b.C5Utpoper == "B" ? "CONSIGNACAO" : b.C5Utpoper == "X" ? "DEMONSTRACAO" : b.C5Utpoper == "F" ? "FATURAMENTO" : "OUTROS",
                                 NF = SD20.D2Doc,
                                 EmissaoNF = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                                 QTD = SD20.D2Quant,
                                 Pedido = b.C5Num,
                                 Cliente = b.C5Nomcli,
                                 Patrimonio = c.C6Upatrim
                             }).ToList();

                Relatorios = query.OrderBy(x => x.EmissaoNF).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("ValidadePermanente");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "ValidadePermanente");

                sheet.Cells[1, 1].Value = "Filial";
                sheet.Cells[1, 2].Value = "Pedido";
                sheet.Cells[1, 3].Value = "NF";
                sheet.Cells[1, 4].Value = "Emissão NF";
                sheet.Cells[1, 5].Value = "Operações";
                sheet.Cells[1, 6].Value = "Agendamento";
                sheet.Cells[1, 7].Value = "Patrimonio";
                sheet.Cells[1, 8].Value = "Produto";
                sheet.Cells[1, 9].Value = "Lote";
                sheet.Cells[1, 10].Value = "QTD";
                sheet.Cells[1, 11].Value = "Validade";
                sheet.Cells[1, 12].Value = "Cliente";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.Pedido;
                    sheet.Cells[i, 3].Value = Pedido.NF;
                    sheet.Cells[i, 4].Value = Pedido.EmissaoNF;
                    sheet.Cells[i, 5].Value = Pedido.Operacoes;
                    sheet.Cells[i, 6].Value = Pedido.Agendamento;
                    sheet.Cells[i, 7].Value = Pedido.Patrimonio;
                    sheet.Cells[i, 8].Value = Pedido.Produto;
                    sheet.Cells[i, 9].Value = Pedido.Lote;
                    sheet.Cells[i, 10].Value = Pedido.QTD;
                    sheet.Cells[i, 11].Value = Pedido.Validade;
                    sheet.Cells[i, 12].Value = Pedido.Cliente;

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
