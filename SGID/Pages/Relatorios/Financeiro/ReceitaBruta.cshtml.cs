using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.AspNetCore.Http.Metadata;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Controladoria;
using SGID.Models.Denuo;
using SGID.Models.Financeiro;
using System.Text.RegularExpressions;

namespace SGID.Pages.Relatorios.Financeiro
{
    public class ReceitaBrutaModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public List<ReceitaBruta> Relatorios { get; set; } = new List<ReceitaBruta> ();
        public ReceitaBrutaModel(TOTVSDENUOContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPost(DateTime DataInicio,DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                var CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };
                var CfNe = new int[] {1202, 1553, 2202, 2553 };

                Relatorios = (from SF20 in Protheus.Sf2010s
                              join SD20 in Protheus.Sd2010s on SF20.F2Doc equals SD20.D2Doc
                              join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                              join SA10 in Protheus.Sa1010s on new { Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja } into Sr
                              from m in Sr.DefaultIfEmpty()
                              join SC50 in Protheus.Sc5010s on SD20.D2Pedido equals SC50.C5Num into Se
                              from c in Se.DefaultIfEmpty()
                              where (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                              && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                              && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                              ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf))
                              && SF20.DELET != "*" && SD20.DELET != "*" && SD20.D2Quant != 0
                              select new ReceitaBruta
                              {
                                  Produto = SB10.B1Cod,
                                  Descricao = SB10.B1Desc,
                                  NFSaida = SF20.F2Doc,
                                  CFOP = SD20.D2Cf,
                                  Cliente = m.A1Nome,
                                  Tipo = m.A1Clinter,
                                  Bruta = SD20.D2Valbrut,
                                  Qtde = SD20.D2Quant,
                                  Parcelas = c.C5Descond.Contains("ATO") && c.C5Descond.Contains("DIAS") ? "1" : c.C5Descond.Contains("ATO") ? c.C5Descond.Substring(6, 2).Replace(" ", "") : c.C5Descond.Contains("PARCELAS") ? c.C5Descond.Substring(0, 2).Replace(" ", "") : "1",
                                  CondPag = c.C5Condpag,
                                  Imposto = SD20.D2Valipi + SD20.D2Valicm + SD20.D2Valimp5 + SD20.D2Valimp6,
                                  Custo = SD20.D2Custo1,
                                  Data = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                                  Pedido = SD20.D2Pedido,
                                  Vendedor = c.C5Nomvend,
                                  Fabricante = SB10.B1Fabric,
                                  UnidadeNegocio = SB10.B1Xunnego == "000001" ? "OUTROS" : SB10.B1Xunnego == "000002" ? "TORAX" : SB10.B1Xunnego == "000003" ? "COLUNA" : SB10.B1Xunnego == "000004" ? "NEURO" : SB10.B1Xunnego == "000005" ? "ORTOPEDIA" : SB10.B1Xunnego == "000006" ? "BUCOMAXILO" : SB10.B1Xunnego == "000007" ? "MATRIX" : SB10.B1Xunnego == "000008" ? "DENTAL" : SB10.B1Xunnego == "000009" ? "CMF" : ""
                              }).ToList();

                var teste = (from SD10 in Protheus.Sd1010s
                             join SF20 in Protheus.Sf2010s on SD10.D1Nfori equals SF20.F2Doc into Sm
                             from c in Sm.DefaultIfEmpty()
                             join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                             join SA10 in Protheus.Sa1010s on c.F2Cliente equals SA10.A1Cod into Sr
                             from m in Sr.DefaultIfEmpty()
                             where (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                             && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             && CfNe.Contains((int)(object)SD10.D1Cf)
                             && c.DELET != "*" && SD10.DELET != "*"
                             select new ReceitaBruta
                             {
                                 Produto = SB10.B1Cod,
                                 Descricao = SB10.B1Desc,
                                 NFSaida = SD10.D1Doc,
                                 CFOP = SD10.D1Cf,
                                 Cliente = m.A1Nome,
                                 Tipo = m.A1Clinter,
                                 Bruta = -(SD10.D1Total - SD10.D1Valdesc + SD10.D1Valipi + SD10.D1Despesa),
                                 Qtde = SD10.D1Quant,
                                 Parcelas = "1",
                                 CondPag = "",
                                 Imposto = -SD10.D1Valipi - SD10.D1Valicm - SD10.D1Valimp5 - SD10.D1Valimp6,
                                 Custo = -SD10.D1Custo,
                                 Data = $"{SD10.D1Emissao.Substring(6, 2)}/{SD10.D1Emissao.Substring(4, 2)}/{SD10.D1Emissao.Substring(0, 4)}",
                                 Pedido = SD10.D1Pedido,
                                 Vendedor = "",
                                 Fabricante = SB10.B1Fabric,
                                 UnidadeNegocio = SB10.B1Xunnego == "000001" ? "OUTROS" : SB10.B1Xunnego == "000002" ? "TORAX" : SB10.B1Xunnego == "000003" ? "COLUNA" : SB10.B1Xunnego == "000004" ? "NEURO" : SB10.B1Xunnego == "000005" ? "ORTOPEDIA" : SB10.B1Xunnego == "000006" ? "BUCOMAXILO" : SB10.B1Xunnego == "000007" ? "MATRIX" : SB10.B1Xunnego == "000008" ? "DENTAL" : SB10.B1Xunnego == "000009" ? "CMF" : ""
                             
                             }).ToList();

                Relatorios.AddRange(teste);

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Receita Bruta", user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim,decimal Selic)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;

                var CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };
                var CfNe = new int[] { 1202, 1553, 2202, 2553, 3202 };

                Relatorios = (from SF20 in Protheus.Sf2010s
                              join SD20 in Protheus.Sd2010s on SF20.F2Doc equals SD20.D2Doc
                              join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                              join SA10 in Protheus.Sa1010s on new { Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja } into Sr
                              from m in Sr.DefaultIfEmpty()
                              join SC50 in Protheus.Sc5010s on SD20.D2Pedido equals SC50.C5Num into Se
                              from c in Se.DefaultIfEmpty()
                              where (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                              && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                              && (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || 
                              ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114) ||
                              ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || 
                              CF.Contains((int)(object)SD20.D2Cf))
                              && SF20.DELET != "*" && SD20.DELET != "*" && SD20.D2Quant != 0
                              select new ReceitaBruta
                              {
                                  Produto = SB10.B1Cod,
                                  Descricao = SB10.B1Desc,
                                  NFSaida = SF20.F2Doc,
                                  CFOP = SD20.D2Cf,
                                  Cliente = m.A1Nome,
                                  Tipo = m.A1Clinter,
                                  Bruta = SD20.D2Valbrut,
                                  Qtde = SD20.D2Quant,
                                  Parcelas = c.C5Descond.Contains("ATO") && c.C5Descond.Contains("DIAS") ? "1": c.C5Descond.Contains("ATO") && c.C5Descond.Contains(",")? "1": c.C5Descond.Contains("ATO") ? c.C5Descond.Substring(6, 2).Replace(" ", "") : c.C5Descond.Contains("PARCELAS") ? c.C5Descond.Substring(0, 2).Replace(" ", "") : "1",
                                  CondPag = c.C5Condpag,
                                  Imposto = SD20.D2Valipi + SD20.D2Valicm + SD20.D2Valimp5 + SD20.D2Valimp6,
                                  Custo = SD20.D2Custo1,
                                  Data = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                                  Pedido = SD20.D2Pedido,
                                  Vendedor = c.C5Nomvend,
                                  Fabricante = SB10.B1Fabric,
                                  UnidadeNegocio = SB10.B1Xunnego == "000001" ? "OUTROS" : SB10.B1Xunnego == "000002" ? "TORAX" : SB10.B1Xunnego == "000003" ? "COLUNA" : SB10.B1Xunnego == "000004" ? "NEURO" : SB10.B1Xunnego == "000005" ? "ORTOPEDIA" : SB10.B1Xunnego == "000006" ? "BUCOMAXILO" : SB10.B1Xunnego == "000007" ? "MATRIX" : SB10.B1Xunnego == "000008" ? "DENTAL" : SB10.B1Xunnego == "000009" ? "CMF" : ""
                              }).ToList();

                var teste = (from SD10 in Protheus.Sd1010s
                             join SF20 in Protheus.Sf2010s on SD10.D1Nfori equals SF20.F2Doc into Sm
                             from c in Sm.DefaultIfEmpty()
                             join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                             join SA10 in Protheus.Sa1010s on c.F2Cliente equals SA10.A1Cod into Sr
                             from m in Sr.DefaultIfEmpty()
                             where (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                             && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             && CfNe.Contains((int)(object)SD10.D1Cf)
                             && c.DELET != "*" && SD10.DELET != "*"
                             select new ReceitaBruta
                             {
                                 Produto = SB10.B1Cod,
                                 Descricao = SB10.B1Desc,
                                 NFSaida = SD10.D1Doc,
                                 CFOP = SD10.D1Cf,
                                 Cliente = m.A1Nome,
                                 Tipo = m.A1Clinter,
                                 Bruta = -(SD10.D1Total - SD10.D1Valdesc + SD10.D1Valipi + SD10.D1Despesa),
                                 Qtde = SD10.D1Quant,
                                 Parcelas = "1",
                                 CondPag = "",
                                 Imposto = -SD10.D1Valipi - SD10.D1Valicm - SD10.D1Valimp5 - SD10.D1Valimp6,
                                 Custo = -SD10.D1Custo,
                                 Data = $"{SD10.D1Emissao.Substring(6, 2)}/{SD10.D1Emissao.Substring(4, 2)}/{SD10.D1Emissao.Substring(0, 4)}",
                                 Pedido = SD10.D1Pedido,
                                 Vendedor = "",
                                 Fabricante = SB10.B1Fabric,
                                 UnidadeNegocio = SB10.B1Xunnego == "000001" ? "OUTROS" : SB10.B1Xunnego == "000002" ? "TORAX" : SB10.B1Xunnego == "000003" ? "COLUNA" : SB10.B1Xunnego == "000004" ? "NEURO" : SB10.B1Xunnego == "000005" ? "ORTOPEDIA" : SB10.B1Xunnego == "000006" ? "BUCOMAXILO" : SB10.B1Xunnego == "000007" ? "MATRIX" : SB10.B1Xunnego == "000008" ? "DENTAL" : SB10.B1Xunnego == "000009" ? "CMF" : ""
                             }).ToList();

                Relatorios.AddRange(teste);


                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("ReceitaBruta");

                var teste2 = Relatorios.Where(x => x.Parcelas == "VE").ToList();

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "ReceitaBruta");

                sheet.Cells[1, 1].Value = "PRODUTO";
                sheet.Cells[1, 2].Value = "DESCRIÇÃO";
                sheet.Cells[1, 3].Value = "NF SAIDA";
                sheet.Cells[1, 4].Value = "CFOP";
                sheet.Cells[1, 5].Value = "CLIENTE";
                sheet.Cells[1, 6].Value = "TIPO INTERMEDIC";
                sheet.Cells[1, 7].Value = "QUANTIDADE";
                sheet.Cells[1, 8].Value = "RECEITA BRUTA";
                sheet.Cells[1, 9].Value = "PARCELAS";
                sheet.Cells[1, 10].Value = "CONDPAG";
                sheet.Cells[1, 11].Value = "TAXA SELIC";
                sheet.Cells[1, 12].Value = "VP - RECEITA BRUTA";
                sheet.Cells[1, 13].Value = "IMPOSTOS S/ NOTAS SAIDAS";
                sheet.Cells[1, 14].Value = "RECEITA LIQUIDA";
                sheet.Cells[1, 15].Value = "CUSTO(CMV)";
                sheet.Cells[1, 16].Value = "CUSTO FINANCEIRO";
                sheet.Cells[1, 17].Value = "MARGEM";
                sheet.Cells[1, 18].Value = "DATA EMISSÃO";
                sheet.Cells[1, 19].Value = "PEDIDO";
                sheet.Cells[1, 20].Value = "VENDEDOR";
                sheet.Cells[1, 21].Value = "FABRICANTE";
                sheet.Cells[1, 22].Value = "UN. NEGOCIO";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Produto;
                    sheet.Cells[i, 2].Value = Pedido.Descricao;
                    sheet.Cells[i, 3].Value = Pedido.NFSaida;
                    sheet.Cells[i, 4].Value = Pedido.CFOP;
                    sheet.Cells[i, 5].Value = Pedido.Cliente;
                    sheet.Cells[i, 6].Value = Pedido.Tipo;

                    sheet.Cells[i, 7].Value = Pedido.Qtde;

                    sheet.Cells[i, 8].Value = Pedido.Bruta;
                    sheet.Cells[i, 8].Style.Numberformat.Format = "0.00";


                    sheet.Cells[i, 9].Value = Convert.ToInt32(Pedido.Parcelas);
                    sheet.Cells[i, 10].Value = Pedido.CondPag;
                    
                    sheet.Cells[i,11].Value = Selic/100;
                    sheet.Cells[i, 11].Style.Numberformat.Format = "0.0000%";

                    sheet.Cells[i, 12].Formula = $"PV(K{i},I{i},-H{i}/I{i},0)";
                    sheet.Cells[i, 12].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 13].Value = Pedido.Imposto;
                    sheet.Cells[i, 13].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 14].Formula = $"+H{i}-M{i}";
                    sheet.Cells[i, 14].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 15].Value = Pedido.Custo;
                    sheet.Cells[i, 15].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 16].Formula = $"+H{i}-L{i}";
                    sheet.Cells[i, 16].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 17].Formula = $"IFERROR((N{i}-O{i}-P{i})/N{i},0)";
                    sheet.Cells[i, 17].Style.Numberformat.Format = "0.00%";

                    sheet.Cells[i, 18].Value = Pedido.Data;

                    sheet.Cells[i, 19].Value = Pedido.Pedido;
                    sheet.Cells[i, 20].Value = Pedido.Vendedor;
                    sheet.Cells[i, 21].Value = Pedido.Fabricante;
                    sheet.Cells[i, 22].Value = Pedido.UnidadeNegocio;
                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CheckMargem.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Receita Bruta Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
