using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Financeiro;
using System.Globalization;

namespace SGID.Pages.Relatorios.Financeiro
{
    [Authorize]
    public class APagarSemanaInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioAPagarMes> Relatorios { get; set; } = new List<RelatorioAPagarMes>();

        public APagarSemanaInterModel(TOTVSINTERContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public string AnoMes { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string AnoMes)
        {
            try
            {
                this.AnoMes = AnoMes;

                var data = DateTime.Now.ToString("yyyy/MM/dd").Replace("/", "");

                var DTMoeda = Protheus.Sm2010s.FirstOrDefault(x => x.DELET != "*" && (int)(object)x.M2Data <= (int)(object)data && x.M2Moeda2 != 0 && x.M2Moeda3 != 0)?.M2Data;

                var Moedas = Protheus.Sm2010s.Where(x => x.DELET != "*" && x.M2Data == DTMoeda).Select(x => new { TxDolar = x.M2Moeda2, TxEuro = x.M2Moeda3 }).FirstOrDefault();

                Relatorios = (from SE20 in Protheus.Se2010s
                              join SED10 in Protheus.Sed010s on SE20.E2Naturez equals SED10.EdCodigo
                              join SX50 in Protheus.Sx5010s on new { Chave = SED10.EdXgrdesp, Filial = "01", Tabela = "Z9" } equals new { Chave = SX50.X5Chave, Filial = SX50.X5Filial, Tabela = SX50.X5Tabela }
                              where SE20.DELET != "*" && SED10.DELET != "*" && SX50.DELET != "*" && SE20.E2Saldo != 0 && SE20.E2Tipo != "PA"
                              && SE20.E2Naturez.Substring(0, 3) != "111" && !(SE20.E2Tipo == "PR" && SE20.E2Prefixo == "EIC")
                              && (int)(object)SE20.E2Vencrea >= (int)(object)$"{AnoMes}01" && (int)(object)SE20.E2Vencrea <= (int)(object)$"{AnoMes}31"
                              select new
                              {
                                  Empresa = "2",
                                  NumDia = SE20.E2Vencrea,
                                  DiaSem = SE20.E2Vencrea,
                                  Semana = SE20.E2Vencrea,
                                  AnoMes = SE20.E2Vencrea.Substring(0, 6),
                                  SE20.E2Naturez,
                                  SED10.EdDescric,
                                  SED10.EdXgrdesp,
                                  GRPDESC = SX50.X5Descri ?? "GRUPO NÃO DEFINIDO",
                                  SE20.E2Fornece,
                                  SE20.E2Loja,
                                  SE20.E2Nomfor,
                                  SE20.E2Vencrea,
                                  SE20.E2Hist,
                                  VLPagar = SE20.E2Tipo == "PRE" ? 0 : SE20.E2Moeda == 1 ? SE20.E2Saldo + SE20.E2Sdacres - SE20.E2Sddecre : SE20.E2Moeda == 2 ? SE20.E2Valor * Moedas.TxDolar : SE20.E2Moeda == 3 ? SE20.E2Valor * Moedas.TxEuro : 0,
                                  VLPrPagar = SE20.E2Tipo != "PRE" ? 0 : SE20.E2Moeda == 1 ? SE20.E2Saldo + SE20.E2Sdacres - SE20.E2Sddecre : SE20.E2Moeda == 2 ? SE20.E2Valor * Moedas.TxDolar : SE20.E2Moeda == 3 ? SE20.E2Valor * Moedas.TxEuro : 0,
                                  VlOrig = SE20.E2Moeda != 1 ? SE20.E2Valor * SE20.E2Txmoeda : SE20.E2Valor,
                                  Tipo = SE20.E2Tipo
                              }).GroupBy(x => new
                              {
                                  x.AnoMes,
                                  x.E2Vencrea,
                                  x.E2Hist,
                                  x.E2Naturez,
                                  x.EdDescric,
                                  x.EdXgrdesp,
                                  x.GRPDESC,
                                  x.E2Fornece,
                                  x.E2Loja,
                                  x.E2Nomfor,
                                  x.Tipo,
                                  x.Empresa,
                                  x.Semana,
                                  x.NumDia,
                                  x.DiaSem
                              }).Select(x => new RelatorioAPagarMes
                              {
                                  Empresa = x.Key.Empresa,
                                  NumDia = x.Key.NumDia,
                                  DiaSem = x.Key.DiaSem,
                                  Semana = x.Key.Semana,
                                  AnoMes = x.Key.AnoMes,
                                  E2Naturez = x.Key.E2Naturez,
                                  EdDescric = x.Key.EdDescric,
                                  EdXgrdesp = x.Key.EdXgrdesp,
                                  GRPDESC = x.Key.GRPDESC,
                                  E2Fornece = x.Key.E2Fornece,
                                  E2Loja = x.Key.E2Loja,
                                  E2Nomfor = x.Key.E2Nomfor,
                                  E2Vencrea = x.Key.E2Vencrea,
                                  E2Hist = x.Key.E2Hist.Trim(),
                                  VLPagar = x.Sum(c => c.VLPagar),
                                  VLPrPagar = x.Sum(c => c.VLPrPagar),
                                  VLOrig = x.Sum(c => c.VlOrig),
                                  Tipo = x.Key.Tipo
                              }).OrderBy(x => x.EdXgrdesp).ThenBy(x => x.Semana).ThenBy(x => x.NumDia).ThenBy(x => x.E2Vencrea).ToList();

                Relatorios.ForEach(x =>
                {
                    var teste = (int)DateTime.ParseExact(x.NumDia, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek + 1;
                    x.NumDia = teste.ToString();
                    x.DiaSem = DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 0 ? "DOM" : (int)DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 1 ? "SEG" : (int)DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 2 ? "TER" : (int)DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 3 ? "QUA" : (int)DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 4 ? "QUI" : (int)DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 5 ? "SEX" : "SAB";
                    x.Semana = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.ParseExact(x.Semana, "yyyyMMdd", CultureInfo.InvariantCulture), CalendarWeekRule.FirstDay, DayOfWeek.Sunday).ToString();
                });

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "APagarSemanaInter",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(string AnoMes)
        {
            try
            {
                this.AnoMes = AnoMes;

                var data = DateTime.Now.ToString("yyyy/MM/dd").Replace("/", "");

                var DTMoeda = Protheus.Sm2010s.FirstOrDefault(x => x.DELET != "*" && (int)(object)x.M2Data <= (int)(object)data && x.M2Moeda2 != 0 && x.M2Moeda3 != 0)?.M2Data;

                var Moedas = Protheus.Sm2010s.Where(x => x.DELET != "*" && x.M2Data == DTMoeda).Select(x => new { TxDolar = x.M2Moeda2, TxEuro = x.M2Moeda3 }).FirstOrDefault();

                Relatorios = (from SE20 in Protheus.Se2010s
                              join SED10 in Protheus.Sed010s on SE20.E2Naturez equals SED10.EdCodigo
                              join SX50 in Protheus.Sx5010s on new { Chave = SED10.EdXgrdesp, Filial = "01", Tabela = "Z9" } equals new { Chave = SX50.X5Chave, Filial = SX50.X5Filial, Tabela = SX50.X5Tabela }
                              where SE20.DELET != "*" && SED10.DELET != "*" && SX50.DELET != "*" && SE20.E2Saldo != 0 && SE20.E2Tipo != "PA"
                              && SE20.E2Naturez.Substring(0, 3) != "111" && !(SE20.E2Tipo == "PR" && SE20.E2Prefixo == "EIC")
                              && (int)(object)SE20.E2Vencrea >= (int)(object)$"{AnoMes}01" && (int)(object)SE20.E2Vencrea <= (int)(object)$"{AnoMes}31"
                              select new
                              {
                                  Empresa = "2",
                                  NumDia = SE20.E2Vencrea,
                                  DiaSem = SE20.E2Vencrea,
                                  Semana = SE20.E2Vencrea,
                                  AnoMes = SE20.E2Vencrea.Substring(0, 6),
                                  SE20.E2Naturez,
                                  SED10.EdDescric,
                                  SED10.EdXgrdesp,
                                  GRPDESC = SX50.X5Descri ?? "GRUPO NÃO DEFINIDO",
                                  SE20.E2Fornece,
                                  SE20.E2Loja,
                                  SE20.E2Nomfor,
                                  SE20.E2Vencrea,
                                  SE20.E2Hist,
                                  VLPagar = SE20.E2Tipo == "PRE" ? 0 : SE20.E2Moeda == 1 ? SE20.E2Saldo + SE20.E2Sdacres - SE20.E2Sddecre : SE20.E2Moeda == 2 ? SE20.E2Valor * Moedas.TxDolar : SE20.E2Moeda == 3 ? SE20.E2Valor * Moedas.TxEuro : 0,
                                  VLPrPagar = SE20.E2Tipo != "PRE" ? 0 : SE20.E2Moeda == 1 ? SE20.E2Saldo + SE20.E2Sdacres - SE20.E2Sddecre : SE20.E2Moeda == 2 ? SE20.E2Valor * Moedas.TxDolar : SE20.E2Moeda == 3 ? SE20.E2Valor * Moedas.TxEuro : 0,
                                  VlOrig = SE20.E2Moeda != 1 ? SE20.E2Valor * SE20.E2Txmoeda : SE20.E2Valor,
                                  Tipo = SE20.E2Tipo
                              }).GroupBy(x => new
                              {
                                  x.AnoMes,
                                  x.E2Vencrea,
                                  x.E2Hist,
                                  x.E2Naturez,
                                  x.EdDescric,
                                  x.EdXgrdesp,
                                  x.GRPDESC,
                                  x.E2Fornece,
                                  x.E2Loja,
                                  x.E2Nomfor,
                                  x.Tipo,
                                  x.Empresa,
                                  x.Semana,
                                  x.NumDia,
                                  x.DiaSem
                              }).Select(x => new RelatorioAPagarMes
                              {
                                  Empresa = x.Key.Empresa,
                                  NumDia = x.Key.NumDia,
                                  DiaSem = x.Key.DiaSem,
                                  Semana = x.Key.Semana,
                                  AnoMes = x.Key.AnoMes,
                                  E2Naturez = x.Key.E2Naturez,
                                  EdDescric = x.Key.EdDescric,
                                  EdXgrdesp = x.Key.EdXgrdesp,
                                  GRPDESC = x.Key.GRPDESC,
                                  E2Fornece = x.Key.E2Fornece,
                                  E2Loja = x.Key.E2Loja,
                                  E2Nomfor = x.Key.E2Nomfor,
                                  E2Vencrea = x.Key.E2Vencrea,
                                  E2Hist = x.Key.E2Hist.Trim(),
                                  VLPagar = x.Sum(c => c.VLPagar),
                                  VLPrPagar = x.Sum(c => c.VLPrPagar),
                                  VLOrig = x.Sum(c => c.VlOrig),
                                  Tipo = x.Key.Tipo
                              }).OrderBy(x => x.EdXgrdesp).ThenBy(x => x.Semana).ThenBy(x => x.NumDia).ThenBy(x => x.E2Vencrea).ToList();

                Relatorios.ForEach(x =>
                {
                    var teste = (int)DateTime.ParseExact(x.NumDia, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek + 1;
                    x.NumDia = teste.ToString();
                    x.DiaSem = DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 0 ? "DOM" : (int)DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 1 ? "SEG" : (int)DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 2 ? "TER" : (int)DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 3 ? "QUA" : (int)DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 4 ? "QUI" : (int)DateTime.ParseExact(x.DiaSem, "yyyyMMdd", CultureInfo.InvariantCulture).DayOfWeek == 5 ? "SEX" : "SAB";
                    x.Semana = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.ParseExact(x.Semana, "yyyyMMdd", CultureInfo.InvariantCulture), CalendarWeekRule.FirstDay, DayOfWeek.Sunday).ToString();
                });

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("APagarSemana");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "APagarSemana");

                sheet.Cells[1, 1].Value = "Grupo de Despesa";
                sheet.Cells[1, 2].Value = "Fornecedor";
                sheet.Cells[1, 3].Value = "Detalhe";
                sheet.Cells[1, 4].Value = "Real";
                sheet.Cells[1, 5].Value = "Provisão";
                sheet.Cells[1, 6].Value = "Total";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.GRPDESC;
                    sheet.Cells[i, 2].Value = Pedido.E2Nomfor;
                    sheet.Cells[i, 3].Value = Pedido.E2Hist;
                    sheet.Cells[i, 4].Value = Pedido.VLPagar;
                    sheet.Cells[i, 5].Value = Pedido.VLPrPagar;
                    sheet.Cells[i, 6].Value = Pedido.VLPagar + Pedido.VLPrPagar;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "APagarSemanaInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "APagarSemanaInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
