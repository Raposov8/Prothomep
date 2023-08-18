using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Financeiro;
using SGID.Models.Relatorio;
using System;
using System.Globalization;

namespace SGID.Pages.Relatorios.Financeiro
{
    [Authorize]
    public class PagoMesModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public List<PagoMes> Relatorios { get; set; } = new List<PagoMes>();
        public PagoMesModel(TOTVSINTERContext protheus, ApplicationDbContext sGID)
        {
            Protheus = protheus;
            SGID = sGID;
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

                var tipo = new string[] { "PA", "NDF" };
                var tipodoc = new string[] { "JR", "MT", "DC", "CM", "D2", "J2", "M2", "V2", "TL", "CP", "C2", "E2", "TR" };

                Relatorios = (from SE50 in Protheus.Se5010s
                              join SED01 in Protheus.Sed010s on SE50.E5Naturez equals SED01.EdCodigo
                              join SX50 in Protheus.Sx5010s on new { Tabela = "Z9", Chave = SED01.EdXgrdesp } equals new { Tabela = SX50.X5Tabela, Chave = SX50.X5Chave }
                              join SA20 in Protheus.Sa2010s on new { Cod = SE50.E5Fornece, Loja = SE50.E5Loja } equals new { Cod = SA20.A2Cod, Loja = SA20.A2Loja }
                              where SE50.DELET != "*" && SED01.DELET != "*" && SX50.DELET != "*" && SA20.DELET != "*"
                              && SE50.E5Recpag == "P" && SE50.E5Tipodoc != "ES" && (SE50.E5Motbx != "CMP" || (tipo.Contains(SE50.E5Tipo) && SE50.E5Motbx == "CMP"))
                              && SE50.E5Motbx != "DAC" && SE50.E5Situaca != "C" && !tipodoc.Contains(SE50.E5Tipodoc)
                              && (int)(object)SE50.E5Data >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SE50.E5Data <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                              && !Protheus.Se5010s.Any(x => x.DELET != "*" && x.E5Filial == SE50.E5Filial && x.E5Prefixo == SE50.E5Prefixo && x.E5Numero == SE50.E5Numero
                              && x.E5Parcela == SE50.E5Parcela && x.E5Tipo == SE50.E5Tipo && x.E5Clifor == SE50.E5Clifor && x.E5Loja == SE50.E5Loja
                              && x.E5Seq == SE50.E5Seq && x.E5Tipodoc == "ES")
                              select new PagoMes
                              {
                                  Empresa = "1",
                                  Numero = SE50.E5Numero,
                                  Parcela = SE50.E5Parcela,
                                  AnoMes = SE50.E5Data.Substring(0, 6),
                                  SubGrupo = SE50.E5Naturez == "211001" || SE50.E5Naturez == "313012" ? "FORNECEDORES NACIONAIS" : SED01.EdXgrdesp == "000001" ? "FORNECEDORES INTERNACIONAIS" : SE50.E5Naturez.Trim() == "311011" ? "BENEFICIOS" : SE50.E5Naturez.Substring(0, 3) == "311" ? "SALARIO" : SE50.E5Naturez.Substring(0, 3) == "312" ? "BENEFICIOS" : SE50.E5Naturez.Substring(0, 3) == "411" ? "ENCARGOS" : SE50.E5Naturez == "IRF" ? "ENCARGOS" : "N/D",
                                  Descric = SED01.EdDescric,
                                  XgrDesp = SED01.EdXgrdesp,
                                  GRPDESC = SX50.X5Descri ?? "GRUPO NAO DEFINIDO",
                                  Fornece = SE50.E5Fornece,
                                  Nome = SA20.A2Nome,
                                  Loja = SE50.E5Loja,
                                  TIT = SE50.E5Histor.Trim(),
                                  Data = SE50.E5Data,
                                  Pago = SE50.E5Valor,
                                  Situacao = SE50.E5Situaca,
                                  Tipo = SE50.E5Tipo,
                                  MotBX = SE50.E5Motbx,
                                  TipoDoc = SE50.E5Tipodoc
                              }
                             ).GroupBy(x => new
                             {
                                 x.Numero,
                                 x.Parcela,
                                 x.AnoMes,
                                 x.Data,
                                 x.SubGrupo,
                                 x.Descric,
                                 x.XgrDesp,
                                 x.GRPDESC,
                                 x.Fornece,
                                 x.Loja,
                                 x.TIT,
                                 x.Situacao,
                                 x.Tipo,
                                 x.MotBX,
                                 x.TipoDoc,
                                 x.Nome
                             }).Select(x => new PagoMes
                             {
                                 Nome = x.Key.Nome,
                                 Numero = x.Key.Numero,
                                 Parcela = x.Key.Parcela,
                                 AnoMes = x.Key.AnoMes,
                                 SubGrupo = x.Key.SubGrupo,
                                 Descric = x.Key.Descric,
                                 XgrDesp = x.Key.XgrDesp,
                                 GRPDESC = x.Key.GRPDESC,
                                 Fornece = x.Key.Fornece,
                                 Loja = x.Key.Loja,
                                 TIT = x.Key.TIT,
                                 Data = x.Key.Data,
                                 Pago = x.Sum(x => x.Pago),
                                 Situacao = x.Key.Situacao,
                                 Tipo = x.Key.Tipo,
                                 MotBX = x.Key.MotBX,
                                 TipoDoc = x.Key.TipoDoc
                             }).OrderBy(x => x.GRPDESC).ThenBy(x => x.SubGrupo).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "PagoMes",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                var tipo = new string[] { "PA", "NDF" };
                var tipodoc = new string[] { "JR", "MT", "DC", "CM", "D2", "J2", "M2", "V2", "TL", "CP", "C2", "E2", "TR" };

                Relatorios = (from SE50 in Protheus.Se5010s
                              join SED01 in Protheus.Sed010s on SE50.E5Naturez equals SED01.EdCodigo
                              join SX50 in Protheus.Sx5010s on new { Tabela = "Z9", Chave = SED01.EdXgrdesp } equals new { Tabela = SX50.X5Tabela, Chave = SX50.X5Chave }
                              join SA20 in Protheus.Sa2010s on new { Cod = SE50.E5Fornece, Loja = SE50.E5Loja } equals new { Cod = SA20.A2Cod, Loja = SA20.A2Loja }
                              where SE50.DELET != "*" && SED01.DELET != "*" && SX50.DELET != "*" && SA20.DELET != "*"
                              && SE50.E5Recpag == "P" && SE50.E5Tipodoc != "ES" && (SE50.E5Motbx != "CMP" || (tipo.Contains(SE50.E5Tipo) && SE50.E5Motbx == "CMP"))
                              && SE50.E5Motbx != "DAC" && SE50.E5Situaca != "C" && !tipodoc.Contains(SE50.E5Tipodoc)
                              && (int)(object)SE50.E5Data >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SE50.E5Data <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                              && !Protheus.Se5010s.Any(x => x.DELET != "*" && x.E5Filial == SE50.E5Filial && x.E5Prefixo == SE50.E5Prefixo && x.E5Numero == SE50.E5Numero
                              && x.E5Parcela == SE50.E5Parcela && x.E5Tipo == SE50.E5Tipo && x.E5Clifor == SE50.E5Clifor && x.E5Loja == SE50.E5Loja
                              && x.E5Seq == SE50.E5Seq && x.E5Tipodoc == "ES")
                              select new PagoMes
                              {
                                  Empresa = "1",
                                  Numero = SE50.E5Numero,
                                  Parcela = SE50.E5Parcela,
                                  AnoMes = SE50.E5Data.Substring(0, 6),
                                  SubGrupo = SE50.E5Naturez == "211001" || SE50.E5Naturez == "313012" ? "FORNECEDORES NACIONAIS" : SED01.EdXgrdesp == "000001" ? "FORNECEDORES INTERNACIONAIS" : SE50.E5Naturez.Trim() == "311011" ? "BENEFICIOS" : SE50.E5Naturez.Substring(0, 3) == "311" ? "SALARIO" : SE50.E5Naturez.Substring(0, 3) == "312" ? "BENEFICIOS" : SE50.E5Naturez.Substring(0, 3) == "411" ? "ENCARGOS" : SE50.E5Naturez == "IRF" ? "ENCARGOS" : "N/D",
                                  Descric = SED01.EdDescric,
                                  XgrDesp = SED01.EdXgrdesp,
                                  GRPDESC = SX50.X5Descri ?? "GRUPO NAO DEFINIDO",
                                  Fornece = SE50.E5Fornece,
                                  Nome = SA20.A2Nome,
                                  Loja = SE50.E5Loja,
                                  TIT = SE50.E5Histor.Trim(),
                                  Data = SE50.E5Data,
                                  Pago = SE50.E5Valor,
                                  Situacao = SE50.E5Situaca,
                                  Tipo = SE50.E5Tipo,
                                  MotBX = SE50.E5Motbx,
                                  TipoDoc = SE50.E5Tipodoc
                              }
                             ).GroupBy(x => new
                             {
                                 x.Numero,
                                 x.Parcela,
                                 x.AnoMes,
                                 x.Data,
                                 x.SubGrupo,
                                 x.Descric,
                                 x.XgrDesp,
                                 x.GRPDESC,
                                 x.Fornece,
                                 x.Loja,
                                 x.TIT,
                                 x.Situacao,
                                 x.Tipo,
                                 x.MotBX,
                                 x.TipoDoc,
                                 x.Nome
                             }).Select(x => new PagoMes
                             {
                                 Nome = x.Key.Nome,
                                 Numero = x.Key.Numero,
                                 Parcela = x.Key.Parcela,
                                 AnoMes = x.Key.AnoMes,
                                 SubGrupo = x.Key.SubGrupo,
                                 Descric = x.Key.Descric,
                                 XgrDesp = x.Key.XgrDesp,
                                 GRPDESC = x.Key.GRPDESC,
                                 Fornece = x.Key.Fornece,
                                 Loja = x.Key.Loja,
                                 TIT = x.Key.TIT,
                                 Data = x.Key.Data,
                                 Pago = x.Sum(x => x.Pago),
                                 Situacao = x.Key.Situacao,
                                 Tipo = x.Key.Tipo,
                                 MotBX = x.Key.MotBX,
                                 TipoDoc = x.Key.TipoDoc
                             }).OrderBy(x => x.GRPDESC).ThenBy(x => x.SubGrupo).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("PagoMes");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "PagoMes");

                sheet.Cells[1, 1].Value = "Grupo de Despesa";
                sheet.Cells[1, 2].Value = "SubGrupo";
                sheet.Cells[1, 3].Value = "Fornecedor";
                sheet.Cells[1, 4].Value = "TIT";
                sheet.Cells[1, 5].Value = "Total";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.GRPDESC;
                    sheet.Cells[i, 2].Value = Pedido.SubGrupo;
                    sheet.Cells[i, 3].Value = Pedido.Nome;
                    sheet.Cells[i, 4].Value = Pedido.TIT;
                    sheet.Cells[i, 5].Value = Pedido.Pago;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "PagoMes.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "PagoMes Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
