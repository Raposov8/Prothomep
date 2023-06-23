using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Denuo;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class NCMModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioNCM> Relatorio = new List<RelatorioNCM>();

        public NCMModel(TOTVSDENUOContext context, ApplicationDbContext sgid)
        {
            Protheus = context;
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

                Relatorio = (from SFT in Protheus.Sft010s
                             join SFT2 in Protheus.Sft010s on new { Filial = SFT.FtFilial, Nfori = SFT.FtNfori, Serie = SFT.FtSerori, Cliefor = SFT.FtCliefor, Loja = SFT.FtLoja, Item = SFT.FtItemori, Prod = SFT.FtProduto }
                             equals new { Filial = SFT2.FtFilial, Nfori = SFT2.FtNfiscal, Serie = SFT2.FtSerie, Cliefor = SFT2.FtCliefor, Loja = SFT2.FtLoja, Item = SFT2.FtItem, Prod = SFT2.FtProduto }
                             join SB10 in Protheus.Sb1010s on SFT.FtProduto equals SB10.B1Cod
                             where SFT.DELET != "*" && SFT2.DELET != "*" && SFT.FtTipomov == "E" && SFT2.FtTipomov == "S"
                             && (int)(object)SFT.FtEntrada >= (int)(object)DataInicio.ToString("dd/MM/yyyy").Replace("/","") && (int)(object)SFT.FtEntrada <= (int)(object)DataFim.ToString("dd/MM/yyyy").Replace("/", "")
                             && SFT.FtNfori != "" && SFT.FtPosipi != SFT2.FtPosipi
                             select new RelatorioNCM
                             {
                                 Filial = SFT.FtFilial,
                                 NotaFiscal = SFT.FtNfiscal,
                                 Serie = SFT.FtSerie,
                                 ClienteFor = SFT.FtCliefor,
                                 Loja = SFT.FtLoja,
                                 Emissao = $"{SFT.FtEmissao.Substring(6, 2)}/{SFT.FtEmissao.Substring(4, 2)}/{SFT.FtEmissao.Substring(0, 4)}",
                                 Entrada = SFT.FtEntrada,
                                 Produto = SFT.FtProduto,
                                 Item = SFT.FtItem,
                                 NCMEntrada = SFT.FtPosipi,
                                 NCMSaida = SFT2.FtPosipi,
                                 Tipo = SB10.B1Tipo
                             }
                             ).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NCM",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {

                Relatorio = (from SFT in Protheus.Sft010s
                             join SFT2 in Protheus.Sft010s on new { Filial = SFT.FtFilial, Nfori = SFT.FtNfori, Serie = SFT.FtSerori, Cliefor = SFT.FtCliefor, Loja = SFT.FtLoja, Item = SFT.FtItemori, Prod = SFT.FtProduto }
                             equals new { Filial = SFT2.FtFilial, Nfori = SFT2.FtNfiscal, Serie = SFT2.FtSerie, Cliefor = SFT2.FtCliefor, Loja = SFT2.FtLoja, Item = SFT2.FtItem, Prod = SFT2.FtProduto }
                             join SB10 in Protheus.Sb1010s on SFT.FtProduto equals SB10.B1Cod
                             where SFT.DELET != "*" && SFT2.DELET != "*" && SFT.FtTipomov == "E" && SFT2.FtTipomov == "S"
                             && (int)(object)SFT.FtEntrada >= (int)(object)DataInicio.ToString("dd/MM/yyyy").Replace("/", "") && (int)(object)SFT.FtEntrada <= (int)(object)DataFim.ToString("dd/MM/yyyy").Replace("/", "")
                             && SFT.FtNfori != "" && SFT.FtPosipi != SFT2.FtPosipi
                             select new RelatorioNCM
                             {
                                 Filial = SFT.FtFilial,
                                 NotaFiscal = SFT.FtNfiscal,
                                 Serie = SFT.FtSerie,
                                 ClienteFor = SFT.FtCliefor,
                                 Loja = SFT.FtLoja,
                                 Emissao = $"{SFT.FtEmissao.Substring(6, 2)}/{SFT.FtEmissao.Substring(4, 2)}/{SFT.FtEmissao.Substring(0, 4)}",
                                 Entrada = SFT.FtEntrada,
                                 Produto = SFT.FtProduto,
                                 Item = SFT.FtItem,
                                 NCMEntrada = SFT.FtPosipi,
                                 NCMSaida = SFT2.FtPosipi,
                                 Tipo = SB10.B1Tipo
                             }
                             ).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("NCM");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "NCM");

                sheet.Cells[1, 1].Value = "FILIAL";
                sheet.Cells[1, 2].Value = "NOTA FISCAL";
                sheet.Cells[1, 3].Value = "SERIE";
                sheet.Cells[1, 4].Value = "CLIENTE FOR";
                sheet.Cells[1, 5].Value = "LOJA";
                sheet.Cells[1, 6].Value = "EMISSAO";
                sheet.Cells[1, 7].Value = "ENTRADA";
                sheet.Cells[1, 8].Value = "PRODUTO";
                sheet.Cells[1, 9].Value = "ITEM";
                sheet.Cells[1, 10].Value = "NCM ENTRADA";
                sheet.Cells[1, 11].Value = "NCM SAIDA";
                sheet.Cells[1, 12].Value = "TIPO";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.NotaFiscal;
                    sheet.Cells[i, 3].Value = Pedido.Serie;
                    sheet.Cells[i, 4].Value = Pedido.ClienteFor;
                    sheet.Cells[i, 5].Value = Pedido.Loja;
                    sheet.Cells[i, 6].Value = Pedido.Emissao;
                    sheet.Cells[i, 7].Value = Pedido.Entrada;
                    sheet.Cells[i, 8].Value = Pedido.Produto;
                    sheet.Cells[i, 9].Value = Pedido.Item;
                    sheet.Cells[i, 10].Value = Pedido.NCMEntrada;
                    sheet.Cells[i, 11].Value = Pedido.NCMSaida;
                    sheet.Cells[i, 12].Value = Pedido.Tipo;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "NCM.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NCM Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
