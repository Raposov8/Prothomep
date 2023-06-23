using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;
using SGID.Models.Financeiro;

namespace SGID.Pages.Relatorios.Financeiro
{
    [Authorize]
    public class APagarMesModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public List<PagoMes> Relatorios { get; set; } = new List<PagoMes>();

        public APagarMesModel(TOTVSDENUOContext protheus,ApplicationDbContext sgid)
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

                var data = DateTime.Now.ToString("yyyy/MM/dd").Replace("/", "");

                var DTMoeda = Protheus.Sm2010s.FirstOrDefault(x => x.DELET != "*" && (int)(object)x.M2Data <= (int)(object)data && x.M2Moeda2 != 0 && x.M2Moeda3 != 0)?.M2Data;

                var Moeda = Protheus.Sm2010s.Where(x => x.DELET != "*" && x.M2Data == DTMoeda).Select(x => new { Dolar = x.M2Moeda2, Euro = x.M2Moeda3 }).FirstOrDefault();

                var query = (from SE20 in Protheus.Se2010s
                             join SED10 in Protheus.Sed010s on SE20.E2Naturez equals SED10.EdCodigo
                             join SX50 in Protheus.Sx5010s on new { Tabela = "Z9", Filial = "03", Chave = SED10.EdXgrdesp } equals new { Tabela = SX50.X5Tabela, Filial = SX50.X5Filial, Chave = SX50.X5Chave }
                             where SED10.DELET != "*" && SX50.DELET != "*" && SE20.DELET != "*" && SE20.E2Saldo != 0 && SE20.E2Tipo != "PA"
                             && SE20.E2Naturez.Substring(0, 3) != "111" && !(SE20.E2Tipo == "PR" && SE20.E2Prefixo == "EIC")
                             && (int)(object)SE20.E2Vencrea >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SE20.E2Vencrea <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             select new
                             {
                                 Empresa = "2",
                                 AnoMes = SE20.E2Vencrea.Substring(0, 6),
                                 SubGrupo = SE20.E2Naturez == "211001" || SE20.E2Naturez == "313012" ? "FORNECEDORES NACIONAIS" : SED10.EdXgrdesp == "000001" ? "FORNECEDORES INTERNACIONAIS" : SE20.E2Naturez.Trim() == "311011" ? "BENEFICIOS" : SE20.E2Naturez.Substring(0, 3) == "311" ? "SALARIO" : SE20.E2Naturez.Substring(0, 3) == "312" ? "BENEFICIOS" : SE20.E2Naturez.Substring(0, 3) == "411" ? "ENCARGOS" : SE20.E2Naturez == "IRF" ? "ENCARGOS" : "N/D",
                                 Descric = SED10.EdDescric,
                                 XgreDesp = SED10.EdXgrdesp,
                                 GrpDesc = SX50.X5Descri ?? "GRUPO NAO DEFINIDO",
                                 Fornece = SE20.E2Fornece,
                                 Loja = SE20.E2Loja,
                                 NomFor = SE20.E2Nomfor,
                                 Vencrea = SE20.E2Vencrea,
                                 VLPAGAR = SE20.E2Moeda == 1 ? SE20.E2Saldo + SE20.E2Sdacres - SE20.E2Sddecre : SE20.E2Moeda == 2 ? SE20.E2Valor * Moeda.Dolar : SE20.E2Moeda == 3 ? SE20.E2Valor * Moeda.Euro : 0,
                                 VLORIG = SE20.E2Moeda != 1 ? SE20.E2Valor * SE20.E2Txmoeda : SE20.E2Valor,
                                 Tipo = SE20.E2Tipo
                             }).GroupBy(x => new
                             {
                                 x.AnoMes,
                                 x.Vencrea,
                                 x.SubGrupo,
                                 x.Descric,
                                 x.XgreDesp,
                                 x.GrpDesc,
                                 x.Fornece,
                                 x.Loja,
                                 x.NomFor,
                                 x.Tipo
                             }).Select(x => new PagoMes
                             {
                                 GRPDESC = x.Key.GrpDesc,
                                 SubGrupo = x.Key.SubGrupo,
                                 Nome = x.Key.NomFor,
                                 Pago = x.Sum(c => c.VLPAGAR)
                             }).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "APagarMes",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                var data = DateTime.Now.ToString("yyyy/MM/dd").Replace("/", "");

                var DTMoeda = Protheus.Sm2010s.FirstOrDefault(x => x.DELET != "*" && (int)(object)x.M2Data <= (int)(object)data && x.M2Moeda2 != 0 && x.M2Moeda3 != 0)?.M2Data;

                var Moeda = Protheus.Sm2010s.Where(x => x.DELET != "*" && x.M2Data == DTMoeda).Select(x => new { Dolar = x.M2Moeda2, Euro = x.M2Moeda3 }).FirstOrDefault();

                var query = (from SE20 in Protheus.Se2010s
                             join SED10 in Protheus.Sed010s on SE20.E2Naturez equals SED10.EdCodigo
                             join SX50 in Protheus.Sx5010s on new { Tabela = "Z9", Filial = "03", Chave = SED10.EdXgrdesp } equals new { Tabela = SX50.X5Tabela, Filial = SX50.X5Filial, Chave = SX50.X5Chave }
                             where SED10.DELET != "*" && SX50.DELET != "*" && SE20.DELET != "*" && SE20.E2Saldo != 0 && SE20.E2Tipo != "PA"
                             && SE20.E2Naturez.Substring(0, 3) != "111" && !(SE20.E2Tipo == "PR" && SE20.E2Prefixo == "EIC")
                             && (int)(object)SE20.E2Vencrea >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SE20.E2Vencrea <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             select new
                             {
                                 Empresa = "2",
                                 AnoMes = SE20.E2Vencrea.Substring(0, 6),
                                 SubGrupo = SE20.E2Naturez == "211001" || SE20.E2Naturez == "313012" ? "FORNECEDORES NACIONAIS" : SED10.EdXgrdesp == "000001" ? "FORNECEDORES INTERNACIONAIS" : SE20.E2Naturez.Trim() == "311011" ? "BENEFICIOS" : SE20.E2Naturez.Substring(0, 3) == "311" ? "SALARIO" : SE20.E2Naturez.Substring(0, 3) == "312" ? "BENEFICIOS" : SE20.E2Naturez.Substring(0, 3) == "411" ? "ENCARGOS" : SE20.E2Naturez == "IRF" ? "ENCARGOS" : "N/D",
                                 Descric = SED10.EdDescric,
                                 XgreDesp = SED10.EdXgrdesp,
                                 GrpDesc = SX50.X5Descri ?? "GRUPO NAO DEFINIDO",
                                 Fornece = SE20.E2Fornece,
                                 Loja = SE20.E2Loja,
                                 NomFor = SE20.E2Nomfor,
                                 Vencrea = SE20.E2Vencrea,
                                 VLPAGAR = SE20.E2Moeda == 1 ? SE20.E2Saldo + SE20.E2Sdacres - SE20.E2Sddecre : SE20.E2Moeda == 2 ? SE20.E2Valor * Moeda.Dolar : SE20.E2Moeda == 3 ? SE20.E2Valor * Moeda.Euro : 0,
                                 VLORIG = SE20.E2Moeda != 1 ? SE20.E2Valor * SE20.E2Txmoeda : SE20.E2Valor,
                                 Tipo = SE20.E2Tipo
                             }).GroupBy(x => new
                             {
                                 x.AnoMes,
                                 x.Vencrea,
                                 x.SubGrupo,
                                 x.Descric,
                                 x.XgreDesp,
                                 x.GrpDesc,
                                 x.Fornece,
                                 x.Loja,
                                 x.NomFor,
                                 x.Tipo
                             }).Select(x => new PagoMes
                             {
                                 GRPDESC = x.Key.GrpDesc,
                                 SubGrupo = x.Key.SubGrupo,
                                 Nome = x.Key.NomFor,
                                 Pago = x.Sum(c => c.VLPAGAR)
                             }).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("APagarMes");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "APagarMes");

                sheet.Cells[1, 1].Value = "Grupo de Despesa";
                sheet.Cells[1, 2].Value = "SubGrupo";
                sheet.Cells[1, 3].Value = "Fornecedor";
                sheet.Cells[1, 4].Value = "Total";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.GRPDESC;
                    sheet.Cells[i, 2].Value = Pedido.SubGrupo;
                    sheet.Cells[i, 3].Value = Pedido.Nome;
                    sheet.Cells[i, 4].Value = Pedido.Pago;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "APagarMes.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "APagarMes Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
