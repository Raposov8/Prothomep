using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;
using SGID.Models.Financeiro;

namespace SGID.Pages.Relatorios.RH
{
    public class BaixaSubDistribuidorModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public List<RelatorioAreceberBaixa> Relatorio { get; set; } = new List<RelatorioAreceberBaixa>();

        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public BaixaSubDistribuidorModel(TOTVSDENUOContext denuo, ApplicationDbContext sgid)
        {
            Protheus = denuo;
            SGID = sgid;
        }
        public void OnGet()
        {
        }

        public IActionResult OnPost(DateTime Datainicio, DateTime Datafim)
        {
            try
            {
                Inicio = Datainicio;
                Fim = Datafim;


                Relatorio = (from SE50 in Protheus.Se5010s
                             join SE10 in Protheus.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                             equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                             join SA10 in Protheus.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                             join SC50 in Protheus.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                             from c in sr.DefaultIfEmpty()
                             where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                             && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                             && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                             && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                             && (int)(object)SE50.E5Data >= (int)(object)Datainicio.ToString("yyyy/MM/dd").Replace("/", "")
                             && (int)(object)SE50.E5Data <= (int)(object)Datafim.ToString("yyyy/MM/dd").Replace("/", "")
                             && SA10.A1Clinter == "S" 
                             select new RelatorioAreceberBaixa
                             {
                                 Prefixo = SE50.E5Prefixo,
                                 Numero = SE50.E5Numero,
                                 Parcela = SE50.E5Parcela,
                                 TP = SE50.E5Tipo,
                                 CliFor = SE50.E5Clifor,
                                 NomeFor = SA10.A1Nome,
                                 Naturez = SE50.E5Naturez,
                                 Vencimento = SE10.E1Vencto,
                                 Historico = SE50.E5Histor,
                                 DataBaixa = SE50.E5Data,
                                 ValorOrig = SE10.E1Valor,
                                 JurMulta = SE50.E5Vljuros + SE50.E5Vlmulta,
                                 Correcao = SE50.E5Vlcorre,
                                 Descon = SE50.E5Vldesco,
                                 Abatimento = 0,
                                 Imposto = 0,
                                 ValorAcess = 0,
                                 TotalBaixado = SE50.E5Valor,
                                 Banco = SE50.E5Banco,
                                 DtDigi = SE50.E5Dtdigit,
                                 Mot = SE50.E5Motbx,
                                 Orig = SE50.E5Filorig,
                                 Vendedor = c.C5Nomvend,
                                 TipoCliente = SA10.A1Clinter
                             }
                             ).ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Relatorio BaixaSubDistribuidor Denuo", user);
            }

            return Page();
        }

        public IActionResult OnPostExport(DateTime Datainicio, DateTime Datafim)
        {
            try
            {
                Inicio = Datainicio;
                Fim = Datafim;

                Inicio = Datainicio;
                Fim = Datafim;

                Relatorio = (from SE50 in Protheus.Se5010s
                             join SE10 in Protheus.Se1010s on new { PRE = SE50.E5Prefixo, Num = SE50.E5Numero, Par = SE50.E5Parcela, Tipo = SE50.E5Tipo, Cliente = SE50.E5Cliente, Loja = SE50.E5Loja }
                             equals new { PRE = SE10.E1Prefixo, Num = SE10.E1Num, Par = SE10.E1Parcela, Tipo = SE10.E1Tipo, Cliente = SE10.E1Cliente, Loja = SE10.E1Loja }
                             join SA10 in Protheus.Sa1010s on SE50.E5Cliente equals SA10.A1Cod
                             join SC50 in Protheus.Sc5010s on SE10.E1Pedido equals SC50.C5Num into sr
                             from c in sr.DefaultIfEmpty()
                             where SE50.DELET != "*" && SE10.DELET != "*" && SE50.E5Recpag == "R"
                             && (SE50.E5Tipodoc == "VL" || SE50.E5Tipodoc == "RA")
                             && (SE50.E5Naturez == "111001" || SE50.E5Naturez == "111004" || SE50.E5Naturez == "111006")
                             && (SE50.E5Banco == "001" || SE50.E5Banco == "237" || SE50.E5Banco == "341")
                             && (int)(object)SE50.E5Data >= (int)(object)Datainicio.ToString("yyyy/MM/dd").Replace("/", "")
                             && (int)(object)SE50.E5Data <= (int)(object)Datafim.ToString("yyyy/MM/dd").Replace("/", "")
                             && SA10.A1Clinter == "S"
                             select new RelatorioAreceberBaixa
                             {
                                 Prefixo = SE50.E5Prefixo,
                                 Numero = SE50.E5Numero,
                                 Parcela = SE50.E5Parcela,
                                 TP = SE50.E5Tipo,
                                 CliFor = SE50.E5Clifor,
                                 NomeFor = SA10.A1Nome,
                                 Naturez = SE50.E5Naturez,
                                 Vencimento = SE10.E1Vencto,
                                 Historico = SE50.E5Histor,
                                 DataBaixa = SE50.E5Data,
                                 ValorOrig = SE10.E1Valor,
                                 JurMulta = SE50.E5Vljuros + SE50.E5Vlmulta,
                                 Correcao = SE50.E5Vlcorre,
                                 Descon = SE50.E5Vldesco,
                                 Abatimento = 0,
                                 Imposto = 0,
                                 ValorAcess = 0,
                                 TotalBaixado = SE50.E5Valor,
                                 Banco = SE50.E5Banco,
                                 DtDigi = SE50.E5Dtdigit,
                                 Mot = SE50.E5Motbx,
                                 Orig = SE50.E5Filorig,
                                 Vendedor = c.C5Nomvend,
                                 TipoCliente = SA10.A1Clinter
                             }
                             ).ToList();


                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("AReceberBaixa");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "AReceberBaixa");

                sheet.Cells[1, 1].Value = "PRF";
                sheet.Cells[1, 2].Value = "NUMERO";
                sheet.Cells[1, 3].Value = "PRC";
                sheet.Cells[1, 4].Value = "TP";
                sheet.Cells[1, 5].Value = "CLI/FOR";
                sheet.Cells[1, 6].Value = "NOME CLI/FOR";
                sheet.Cells[1, 7].Value = "NATUREZA";
                sheet.Cells[1, 8].Value = "VENCTO";
                sheet.Cells[1, 9].Value = "HISTORICO";
                sheet.Cells[1, 10].Value = "DT BAIXA";
                sheet.Cells[1, 11].Value = "VALOR ORIGINAL";
                sheet.Cells[1, 12].Value = "JUR/MULTA";
                sheet.Cells[1, 13].Value = "CORRECAO";
                sheet.Cells[1, 14].Value = "DESCONTOS";
                sheet.Cells[1, 15].Value = "ABATIM.";
                sheet.Cells[1, 16].Value = "IMPOSTOS";
                sheet.Cells[1, 17].Value = "VALOR ACESSORIO";
                sheet.Cells[1, 18].Value = "TOTAL BAIXADO";
                sheet.Cells[1, 19].Value = "BCO";
                sheet.Cells[1, 20].Value = "DT DIG";
                sheet.Cells[1, 21].Value = "MOT";
                sheet.Cells[1, 22].Value = "ORIG";
                sheet.Cells[1, 23].Value = "VENDEDOR";
                sheet.Cells[1, 24].Value = "TIPOCLIENTE";


                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Prefixo;
                    sheet.Cells[i, 2].Value = Pedido.Numero;
                    sheet.Cells[i, 3].Value = Pedido.Parcela;
                    sheet.Cells[i, 4].Value = Pedido.TP;
                    sheet.Cells[i, 5].Value = Pedido.CliFor;
                    sheet.Cells[i, 6].Value = Pedido.NomeFor;
                    sheet.Cells[i, 7].Value = Pedido.Naturez;
                    sheet.Cells[i, 8].Value = Pedido.Vencimento;
                    sheet.Cells[i, 9].Value = Pedido.Historico;
                    sheet.Cells[i, 10].Value = Pedido.DataBaixa;
                    sheet.Cells[i, 11].Value = Pedido.ValorOrig;
                    sheet.Cells[i, 11].Style.Numberformat.Format = "0.00";
                    sheet.Cells[i, 12].Value = Pedido.JurMulta;
                    sheet.Cells[i, 12].Style.Numberformat.Format = "0.00";
                    sheet.Cells[i, 13].Value = Pedido.Correcao;
                    sheet.Cells[i, 13].Style.Numberformat.Format = "0.00";
                    sheet.Cells[i, 14].Value = Pedido.Descon;
                    sheet.Cells[i, 14].Style.Numberformat.Format = "0.00";
                    sheet.Cells[i, 15].Value = Pedido.Abatimento;
                    sheet.Cells[i, 15].Style.Numberformat.Format = "0.00";
                    sheet.Cells[i, 16].Value = Pedido.Imposto;
                    sheet.Cells[i, 16].Style.Numberformat.Format = "0.00";
                    sheet.Cells[i, 17].Value = Pedido.ValorAcess;
                    sheet.Cells[i, 17].Style.Numberformat.Format = "0.00";
                    sheet.Cells[i, 18].Value = Pedido.TotalBaixado;
                    sheet.Cells[i, 18].Style.Numberformat.Format = "0.00";
                    sheet.Cells[i, 19].Value = Pedido.Banco;
                    sheet.Cells[i, 20].Value = Pedido.DtDigi;
                    sheet.Cells[i, 21].Value = Pedido.Mot;
                    sheet.Cells[i, 22].Value = Pedido.Orig;
                    sheet.Cells[i, 23].Value = Pedido.Vendedor;
                    sheet.Cells[i, 24].Value = Pedido.TipoCliente;



                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "BaixaSubDistribuidorDenuo.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Relatorio BaixaSubDistribuidor Denuo Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
