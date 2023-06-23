using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OPMEnexo;
using SGID.Data;
using SGID.Data.Model;
using SGID.Data.Models;
using SGID.Models.Controladoria;
using SGID.Models.Denuo;
using SGID.Models.RH;

namespace SGID.Pages.Relatorios.RH
{

    [Authorize]
    public class FaturamentoDentalModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioFaturamentoDental> Relatorio { get; set; } = new List<RelatorioFaturamentoDental>();
        public List<RelatorioDevolucaoDental> Devolucao { get; set; } = new List<RelatorioDevolucaoDental>();
        [BindProperty]
        public DateTime Inicio { get; set; }
        [BindProperty]
        public DateTime Fim { get; set; }
        public List<string> Vendedores { get; set; } = new List<string>();
        public string Vendedor { get;set; } = "";

        public FaturamentoDentalModel(TOTVSDENUOContext context,ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        public void OnGet()
        {
            Vendedores = Protheus.Sa3010s.Where(x => x.A3Xdescun == "DENTAL                        " && x.A3Msblql != "1" && x.DELET != "*").Select(x => x.A3Nome).ToList();
        }

        public IActionResult OnPostAsync(DateTime DataInicio,DateTime DataFim,string vende)
        {

            try
            {
                Inicio = DataInicio;
                Fim = DataFim;
                Vendedor = vende;

                var cf = new[] { "5551", "6551", "6107", "6109" };

                if (vende == "" || vende == null)
                {
                    var query = (from SD20 in Protheus.Sd2010s
                                 join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                 join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                 join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                 join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                                 join SC50 in Protheus.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num, Cli = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                                 join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod into sr
                                 from c in sr.DefaultIfEmpty()
                                 where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*"
                                 && SC60.DELET != "*" && SC50.DELET != "*" && c.DELET != "*" &&
                                 (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114)
                                 || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114)
                                 || ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114)
                                 || cf.Contains(SD20.D2Cf)) && (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                                 && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "") && SD20.D2Quant != 0 && SC50.C5Xtipopv == "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140"
                                 && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                 select new RelatorioFaturamentoDental
                                 {
                                     D2_FILIAL = SD20.D2Filial,
                                     D2_CLIENTE = SD20.D2Cliente,
                                     D2_LOJA = SD20.D2Emissao,
                                     A1_NOME = SA10.A1Nome,
                                     A1_CLINTER = SA10.A1Clinter,
                                     D2_DOC = SD20.D2Doc,
                                     D2_SERIE = SD20.D2Serie,
                                     D2_EMISSAO = SD20.D2Emissao,
                                     D2_PEDIDO = SD20.D2Pedido,
                                     C5_UNUMAGE = SC50.C5Unumage,
                                     C5_EMISSAO = SC50.C5Emissao,
                                     C5_VEND1 = SC50.C5Vend1,
                                     A3_NOME = c.A3Nome,
                                     C5_X_DTCIR = SC50.C5XDtcir,
                                     C5_X_NMMED = SC50.C5XNmmed,
                                     C5_X_NMPAC = SC50.C5XNmpac,
                                     C5_X_NMPLA = SC50.C5XNmpla,
                                     C5_UTPOPER = SC50.C5Utpoper,
                                     C6_PRODUTO = SC60.C6Produto,
                                     B1_DESC = SB10.B1Desc,
                                     D2_TOTAL = SD20.D2Total,
                                     D2_QUANT = SD20.D2Quant,
                                     D2_VALIPI = SD20.D2Valipi,
                                     D2_VALICM = SD20.D2Valicm,
                                     D2_DESCON = SD20.D2Descon
                                 }
                                  ).GroupBy(Prod =>
                                  new
                                  {
                                      Prod.D2_FILIAL,
                                      Prod.D2_CLIENTE,
                                      Prod.D2_LOJA,
                                      Prod.A1_NOME,
                                      Prod.A1_CLINTER,
                                      Prod.D2_DOC,
                                      Prod.D2_SERIE,
                                      Prod.D2_EMISSAO,
                                      Prod.D2_PEDIDO,
                                      Prod.C5_UNUMAGE,
                                      Prod.C5_EMISSAO,
                                      Prod.C5_VEND1,
                                      Prod.A3_NOME,
                                      Prod.C5_X_DTCIR,
                                      Prod.C5_X_NMMED,
                                      Prod.C5_X_NMPAC,
                                      Prod.C5_X_NMPLA,
                                      Prod.C5_UTPOPER,
                                      Prod.C6_PRODUTO,
                                      Prod.B1_DESC
                                  });

                    Relatorio = query.Select(x => new RelatorioFaturamentoDental
                    {
                        D2_FILIAL = x.Key.D2_FILIAL,
                        D2_CLIENTE = x.Key.D2_CLIENTE,
                        D2_LOJA = x.Key.D2_LOJA,
                        A1_NOME = x.Key.A1_NOME,
                        A1_CLINTER = x.Key.A1_CLINTER,
                        D2_DOC = x.Key.D2_DOC,
                        D2_SERIE = x.Key.D2_SERIE,
                        D2_EMISSAO = $"{x.Key.D2_EMISSAO.Substring(6, 2)}/{x.Key.D2_EMISSAO.Substring(4, 2)}/{x.Key.D2_EMISSAO.Substring(0, 4)}",
                        D2_PEDIDO = x.Key.D2_PEDIDO,
                        C5_UNUMAGE = x.Key.C5_UNUMAGE,
                        C5_EMISSAO = x.Key.C5_EMISSAO,
                        C5_VEND1 = x.Key.C5_VEND1,
                        A3_NOME = x.Key.A3_NOME,
                        C5_X_DTCIR = x.Key.C5_X_DTCIR,
                        C5_X_NMMED = x.Key.C5_X_NMMED,
                        C5_X_NMPAC = x.Key.C5_X_NMPAC,
                        C5_X_NMPLA = x.Key.C5_X_NMPLA,
                        C5_UTPOPER = x.Key.C5_UTPOPER,
                        C6_PRODUTO = x.Key.C6_PRODUTO,
                        B1_DESC = x.Key.B1_DESC,
                        D2_TOTAL = x.Sum(c => c.D2_TOTAL),
                        D2_QUANT = x.Sum(c => c.D2_QUANT),
                        D2_VALIPI = x.Sum(c => c.D2_VALIPI),
                        D2_VALICM = x.Sum(c => c.D2_VALICM),
                        D2_DESCON = x.Sum(c => c.D2_DESCON),
                    }).OrderBy(x => x.A3_NOME).ToList();


                    var CfDevolucao = new string[] { "1202", "2202", "3202", "1553", "2553" };

                    var queryDevolucao = (from SD10 in Protheus.Sd1010s
                                          join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Fornece = SF20.F2Cliente, Loja = SF20.F2Loja }
                                          join SD20 in Protheus.Sd2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja, Item = SD10.D1Itemori } equals new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Fornece = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Item }
                                          join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                          join SA10 in Protheus.Sa1010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA10.A1Cod, Loja = SA10.A1Loja }
                                          join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                          where SD10.DELET != "*" && SF20.DELET != "*" && SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SC50.DELET != "*"
                                          && CfDevolucao.Contains(SD10.D1Cf) && SC50.C5Xtipopv == "D"
                                          && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                                          && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                                          select new
                                          {
                                              SD10.D1Filial,
                                              SD10.D1Fornece,
                                              SD10.D1Loja,
                                              SA10.A1Nome,
                                              SA10.A1Clinter,
                                              SD10.D1Doc,
                                              SD10.D1Serie,
                                              SD10.D1Emissao,
                                              SD20.D2Doc,
                                              SD20.D2Emissao,
                                              SD10.D1Total,
                                              SD10.D1Valipi,
                                              SD10.D1Valicm,
                                              SC50.C5Vend1,
                                              SC50.C5Nomvend,
                                              SD10.D1Valdesc,
                                              SD10.D1Dtdigit,
                                              SA10.A1Est,
                                              SA10.A1Mun
                                          })
                                          .GroupBy(x => new
                                          {
                                              x.D1Filial,
                                              x.D1Fornece,
                                              x.D1Loja,
                                              x.A1Nome,
                                              x.A1Clinter,
                                              x.A1Est,
                                              x.A1Mun,
                                              x.D1Doc,
                                              x.D1Serie,
                                              x.D1Emissao,
                                              x.D1Dtdigit,
                                              x.D2Doc,
                                              x.D2Emissao,
                                              x.C5Vend1,
                                              x.C5Nomvend
                                          });

                    Devolucao = queryDevolucao.Select(x => new RelatorioDevolucaoDental
                    {
                        Emissao = $"{x.Key.D1Emissao.Substring(6, 2)}/{x.Key.D1Emissao.Substring(4, 2)}/{x.Key.D1Emissao.Substring(0, 4)}",
                        Filial = x.Key.D1Filial,
                        NumNF = x.Key.D1Doc,
                        EmissaoOrig = $"{x.Key.D2Emissao.Substring(6, 2)}/{x.Key.D2Emissao.Substring(4, 2)}/{x.Key.D2Emissao.Substring(0, 4)}",
                        NFOrig = x.Key.D2Doc,
                        CodVend = x.Key.C5Vend1,
                        Vendedor = x.Key.C5Nomvend,
                        Cliente = x.Key.A1Nome,
                        Total = -(x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valipi)),
                    }).OrderBy(x => x.Vendedor).ToList();
                }
                else
                {
                    var query = (from SD20 in Protheus.Sd2010s
                                 join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                 join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                 join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                 join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                                 join SC50 in Protheus.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num, Cli = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                                 join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod into sr
                                 from c in sr.DefaultIfEmpty()
                                 where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*"
                                 && SC60.DELET != "*" && SC50.DELET != "*" && c.DELET != "*" &&
                                 (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114)
                                 || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114)
                                 || ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114)
                                 || cf.Contains(SD20.D2Cf)) && (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                                 && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "") && SD20.D2Quant != 0 && SC50.C5Xtipopv == "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140"
                                 && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && c.A3Nome == vende
                                 select new RelatorioFaturamentoDental
                                 {
                                     D2_FILIAL = SD20.D2Filial,
                                     D2_CLIENTE = SD20.D2Cliente,
                                     D2_LOJA = SD20.D2Emissao,
                                     A1_NOME = SA10.A1Nome,
                                     A1_CLINTER = SA10.A1Clinter,
                                     D2_DOC = SD20.D2Doc,
                                     D2_SERIE = SD20.D2Serie,
                                     D2_EMISSAO = SD20.D2Emissao,
                                     D2_PEDIDO = SD20.D2Pedido,
                                     C5_UNUMAGE = SC50.C5Unumage,
                                     C5_EMISSAO = SC50.C5Emissao,
                                     C5_VEND1 = SC50.C5Vend1,
                                     A3_NOME = c.A3Nome,
                                     C5_X_DTCIR = SC50.C5XDtcir,
                                     C5_X_NMMED = SC50.C5XNmmed,
                                     C5_X_NMPAC = SC50.C5XNmpac,
                                     C5_X_NMPLA = SC50.C5XNmpla,
                                     C5_UTPOPER = SC50.C5Utpoper,
                                     C6_PRODUTO = SC60.C6Produto,
                                     B1_DESC = SB10.B1Desc,
                                     D2_TOTAL = SD20.D2Total,
                                     D2_QUANT = SD20.D2Quant,
                                     D2_VALIPI = SD20.D2Valipi,
                                     D2_VALICM = SD20.D2Valicm,
                                     D2_DESCON = SD20.D2Descon
                                 }
                                  ).GroupBy(Prod =>
                                  new
                                  {
                                      Prod.D2_FILIAL,
                                      Prod.D2_CLIENTE,
                                      Prod.D2_LOJA,
                                      Prod.A1_NOME,
                                      Prod.A1_CLINTER,
                                      Prod.D2_DOC,
                                      Prod.D2_SERIE,
                                      Prod.D2_EMISSAO,
                                      Prod.D2_PEDIDO,
                                      Prod.C5_UNUMAGE,
                                      Prod.C5_EMISSAO,
                                      Prod.C5_VEND1,
                                      Prod.A3_NOME,
                                      Prod.C5_X_DTCIR,
                                      Prod.C5_X_NMMED,
                                      Prod.C5_X_NMPAC,
                                      Prod.C5_X_NMPLA,
                                      Prod.C5_UTPOPER,
                                      Prod.C6_PRODUTO,
                                      Prod.B1_DESC
                                  });

                    Relatorio = query.Select(x => new RelatorioFaturamentoDental
                    {
                        D2_FILIAL = x.Key.D2_FILIAL,
                        D2_CLIENTE = x.Key.D2_CLIENTE,
                        D2_LOJA = x.Key.D2_LOJA,
                        A1_NOME = x.Key.A1_NOME,
                        A1_CLINTER = x.Key.A1_CLINTER,
                        D2_DOC = x.Key.D2_DOC,
                        D2_SERIE = x.Key.D2_SERIE,
                        D2_EMISSAO = $"{x.Key.D2_EMISSAO.Substring(6, 2)}/{x.Key.D2_EMISSAO.Substring(4, 2)}/{x.Key.D2_EMISSAO.Substring(0, 4)}",
                        D2_PEDIDO = x.Key.D2_PEDIDO,
                        C5_UNUMAGE = x.Key.C5_UNUMAGE,
                        C5_EMISSAO = x.Key.C5_EMISSAO,
                        C5_VEND1 = x.Key.C5_VEND1,
                        A3_NOME = x.Key.A3_NOME,
                        C5_X_DTCIR = x.Key.C5_X_DTCIR,
                        C5_X_NMMED = x.Key.C5_X_NMMED,
                        C5_X_NMPAC = x.Key.C5_X_NMPAC,
                        C5_X_NMPLA = x.Key.C5_X_NMPLA,
                        C5_UTPOPER = x.Key.C5_UTPOPER,
                        C6_PRODUTO = x.Key.C6_PRODUTO,
                        B1_DESC = x.Key.B1_DESC,
                        D2_TOTAL = x.Sum(c => c.D2_TOTAL),
                        D2_QUANT = x.Sum(c => c.D2_QUANT),
                        D2_VALIPI = x.Sum(c => c.D2_VALIPI),
                        D2_VALICM = x.Sum(c => c.D2_VALICM),
                        D2_DESCON = x.Sum(c => c.D2_DESCON),
                    }).OrderBy(x => x.A3_NOME).ToList();


                    var CfDevolucao = new string[] { "1202", "2202", "3202", "1553", "2553" };

                    var queryDevolucao = (from SD10 in Protheus.Sd1010s
                                          join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Fornece = SF20.F2Cliente, Loja = SF20.F2Loja }
                                          join SD20 in Protheus.Sd2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja, Item = SD10.D1Itemori } equals new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Fornece = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Item }
                                          join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                          join SA10 in Protheus.Sa1010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA10.A1Cod, Loja = SA10.A1Loja }
                                          join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                          where SD10.DELET != "*" && SF20.DELET != "*" && SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SC50.DELET != "*"
                                          && CfDevolucao.Contains(SD10.D1Cf) && SC50.C5Xtipopv == "D"
                                          && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                                          && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                                          && SC50.C5Nomvend == vende
                                          select new
                                          {
                                              SD10.D1Filial,
                                              SD10.D1Fornece,
                                              SD10.D1Loja,
                                              SA10.A1Nome,
                                              SA10.A1Clinter,
                                              SD10.D1Doc,
                                              SD10.D1Serie,
                                              SD10.D1Emissao,
                                              SD20.D2Doc,
                                              SD20.D2Emissao,
                                              SD10.D1Total,
                                              SD10.D1Valipi,
                                              SD10.D1Valicm,
                                              SC50.C5Vend1,
                                              SC50.C5Nomvend,
                                              SD10.D1Valdesc,
                                              SD10.D1Dtdigit,
                                              SA10.A1Est,
                                              SA10.A1Mun
                                          })
                                          .GroupBy(x => new
                                          {
                                              x.D1Filial,
                                              x.D1Fornece,
                                              x.D1Loja,
                                              x.A1Nome,
                                              x.A1Clinter,
                                              x.A1Est,
                                              x.A1Mun,
                                              x.D1Doc,
                                              x.D1Serie,
                                              x.D1Emissao,
                                              x.D1Dtdigit,
                                              x.D2Doc,
                                              x.D2Emissao,
                                              x.C5Vend1,
                                              x.C5Nomvend
                                          });

                    Devolucao = queryDevolucao.Select(x => new RelatorioDevolucaoDental
                    {
                        Emissao = $"{x.Key.D1Emissao.Substring(6, 2)}/{x.Key.D1Emissao.Substring(4, 2)}/{x.Key.D1Emissao.Substring(0, 4)}",
                        Filial = x.Key.D1Filial,
                        NumNF = x.Key.D1Doc,
                        EmissaoOrig = $"{x.Key.D2Emissao.Substring(6, 2)}/{x.Key.D2Emissao.Substring(4, 2)}/{x.Key.D2Emissao.Substring(0, 4)}",
                        NFOrig = x.Key.D2Doc,
                        CodVend = x.Key.C5Vend1,
                        Vendedor = x.Key.C5Nomvend,
                        Cliente = x.Key.A1Nome,
                        Total = -(x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valipi)),
                    }).OrderBy(x => x.Vendedor).ToList();
                }

                Vendedores = Protheus.Sa3010s.Where(x => x.A3Xdescun == "DENTAL                        " && x.A3Msblql != "1" && x.DELET != "*").Select(x => x.A3Nome).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoDental",user);

                return LocalRedirect("/error");
            }
        }


        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim, string vende)
        {

            try
            {
                var cf = new[] { "5551", "6551", "6107", "6109" };
                if (vende == "" || vende == null)
                {
                    var query = (from SD20 in Protheus.Sd2010s
                                 join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                 join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                 join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                 join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                                 join SC50 in Protheus.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num, Cli = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                                 join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod into sr
                                 from c in sr.DefaultIfEmpty()
                                 where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*"
                                 && SC60.DELET != "*" && SC50.DELET != "*" && c.DELET != "*" &&
                                 (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114)
                                 || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114)
                                 || ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114)
                                 || cf.Contains(SD20.D2Cf)) && (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                                 && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "") && SD20.D2Quant != 0 && SC50.C5Xtipopv == "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140"
                                 && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140"
                                 select new RelatorioFaturamentoDental
                                 {
                                     D2_FILIAL = SD20.D2Filial,
                                     D2_CLIENTE = SD20.D2Cliente,
                                     D2_LOJA = SD20.D2Emissao,
                                     A1_NOME = SA10.A1Nome,
                                     A1_CLINTER = SA10.A1Clinter,
                                     D2_DOC = SD20.D2Doc,
                                     D2_SERIE = SD20.D2Serie,
                                     D2_EMISSAO = SD20.D2Emissao,
                                     D2_PEDIDO = SD20.D2Pedido,
                                     C5_UNUMAGE = SC50.C5Unumage,
                                     C5_EMISSAO = SC50.C5Emissao,
                                     C5_VEND1 = SC50.C5Vend1,
                                     A3_NOME = c.A3Nome,
                                     C5_X_DTCIR = SC50.C5XDtcir,
                                     C5_X_NMMED = SC50.C5XNmmed,
                                     C5_X_NMPAC = SC50.C5XNmpac,
                                     C5_X_NMPLA = SC50.C5XNmpla,
                                     C5_UTPOPER = SC50.C5Utpoper,
                                     C6_PRODUTO = SC60.C6Produto,
                                     B1_DESC = SB10.B1Desc,
                                     D2_TOTAL = SD20.D2Total,
                                     D2_QUANT = SD20.D2Quant,
                                     D2_VALIPI = SD20.D2Valipi,
                                     D2_VALICM = SD20.D2Valicm,
                                     D2_DESCON = SD20.D2Descon
                                 }
                                  ).GroupBy(Prod =>
                                  new
                                  {
                                      Prod.D2_FILIAL,
                                      Prod.D2_CLIENTE,
                                      Prod.D2_LOJA,
                                      Prod.A1_NOME,
                                      Prod.A1_CLINTER,
                                      Prod.D2_DOC,
                                      Prod.D2_SERIE,
                                      Prod.D2_EMISSAO,
                                      Prod.D2_PEDIDO,
                                      Prod.C5_UNUMAGE,
                                      Prod.C5_EMISSAO,
                                      Prod.C5_VEND1,
                                      Prod.A3_NOME,
                                      Prod.C5_X_DTCIR,
                                      Prod.C5_X_NMMED,
                                      Prod.C5_X_NMPAC,
                                      Prod.C5_X_NMPLA,
                                      Prod.C5_UTPOPER,
                                      Prod.C6_PRODUTO,
                                      Prod.B1_DESC
                                  });

                    Relatorio = query.Select(x => new RelatorioFaturamentoDental
                    {
                        D2_FILIAL = x.Key.D2_FILIAL,
                        D2_CLIENTE = x.Key.D2_CLIENTE,
                        D2_LOJA = x.Key.D2_LOJA,
                        A1_NOME = x.Key.A1_NOME,
                        A1_CLINTER = x.Key.A1_CLINTER,
                        D2_DOC = x.Key.D2_DOC,
                        D2_SERIE = x.Key.D2_SERIE,
                        D2_EMISSAO = $"{x.Key.D2_EMISSAO.Substring(6, 2)}/{x.Key.D2_EMISSAO.Substring(4, 2)}/{x.Key.D2_EMISSAO.Substring(0, 4)}",
                        D2_PEDIDO = x.Key.D2_PEDIDO,
                        C5_UNUMAGE = x.Key.C5_UNUMAGE,
                        C5_EMISSAO = x.Key.C5_EMISSAO,
                        C5_VEND1 = x.Key.C5_VEND1,
                        A3_NOME = x.Key.A3_NOME,
                        C5_X_DTCIR = x.Key.C5_X_DTCIR,
                        C5_X_NMMED = x.Key.C5_X_NMMED,
                        C5_X_NMPAC = x.Key.C5_X_NMPAC,
                        C5_X_NMPLA = x.Key.C5_X_NMPLA,
                        C5_UTPOPER = x.Key.C5_UTPOPER,
                        C6_PRODUTO = x.Key.C6_PRODUTO,
                        B1_DESC = x.Key.B1_DESC,
                        D2_TOTAL = x.Sum(c => c.D2_TOTAL),
                        D2_QUANT = x.Sum(c => c.D2_QUANT),
                        D2_VALIPI = x.Sum(c => c.D2_VALIPI),
                        D2_VALICM = x.Sum(c => c.D2_VALICM),
                        D2_DESCON = x.Sum(c => c.D2_DESCON),
                    }).OrderBy(x => x.A3_NOME).ToList();


                    var CfDevolucao = new string[] { "1202", "2202", "3202", "1553", "2553" };

                    var queryDevolucao = (from SD10 in Protheus.Sd1010s
                                          join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Fornece = SF20.F2Cliente, Loja = SF20.F2Loja }
                                          join SD20 in Protheus.Sd2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja, Item = SD10.D1Itemori } equals new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Fornece = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Item }
                                          join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                          join SA10 in Protheus.Sa1010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA10.A1Cod, Loja = SA10.A1Loja }
                                          join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                          where SD10.DELET != "*" && SF20.DELET != "*" && SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SC50.DELET != "*"
                                          && CfDevolucao.Contains(SD10.D1Cf) && SC50.C5Xtipopv == "D"
                                          && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                                          && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                                          select new
                                          {
                                              SD10.D1Filial,
                                              SD10.D1Fornece,
                                              SD10.D1Loja,
                                              SA10.A1Nome,
                                              SA10.A1Clinter,
                                              SD10.D1Doc,
                                              SD10.D1Serie,
                                              SD10.D1Emissao,
                                              SD20.D2Doc,
                                              SD20.D2Emissao,
                                              SD10.D1Total,
                                              SD10.D1Valipi,
                                              SD10.D1Valicm,
                                              SC50.C5Vend1,
                                              SC50.C5Nomvend,
                                              SD10.D1Valdesc,
                                              SD10.D1Dtdigit,
                                              SA10.A1Est,
                                              SA10.A1Mun
                                          })
                                          .GroupBy(x => new
                                          {
                                              x.D1Filial,
                                              x.D1Fornece,
                                              x.D1Loja,
                                              x.A1Nome,
                                              x.A1Clinter,
                                              x.A1Est,
                                              x.A1Mun,
                                              x.D1Doc,
                                              x.D1Serie,
                                              x.D1Emissao,
                                              x.D1Dtdigit,
                                              x.D2Doc,
                                              x.D2Emissao,
                                              x.C5Vend1,
                                              x.C5Nomvend
                                          });

                    Devolucao = queryDevolucao.Select(x => new RelatorioDevolucaoDental
                    {
                        Emissao = $"{x.Key.D1Emissao.Substring(6, 2)}/{x.Key.D1Emissao.Substring(4, 2)}/{x.Key.D1Emissao.Substring(0, 4)}",
                        Filial = x.Key.D1Filial,
                        NumNF = x.Key.D1Doc,
                        EmissaoOrig = $"{x.Key.D2Emissao.Substring(6, 2)}/{x.Key.D2Emissao.Substring(4, 2)}/{x.Key.D2Emissao.Substring(0, 4)}",
                        NFOrig = x.Key.D2Doc,
                        CodVend = x.Key.C5Vend1,
                        Vendedor = x.Key.C5Nomvend,
                        Cliente = x.Key.A1Nome,
                        Total = -(x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valipi)),
                    }).OrderBy(x => x.Vendedor).ToList();
                }
                else
                {
                    var query = (from SD20 in Protheus.Sd2010s
                                 join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                                 join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                 join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                 join SC60 in Protheus.Sc6010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido, Cli = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Itempv, Cod = SD20.D2Cod } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja, Item = SC60.C6Item, Cod = SC60.C6Produto }
                                 join SC50 in Protheus.Sc5010s on new { Filial = SC60.C6Filial, Num = SC60.C6Num, Cli = SC60.C6Cli, Loja = SC60.C6Loja } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num, Cli = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                                 join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod into sr
                                 from c in sr.DefaultIfEmpty()
                                 where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*"
                                 && SC60.DELET != "*" && SC50.DELET != "*" && c.DELET != "*" &&
                                 (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114)
                                 || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114)
                                 || ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114)
                                 || cf.Contains(SD20.D2Cf)) && (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                                 && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "") && SD20.D2Quant != 0 && SC50.C5Xtipopv == "D" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140"
                                 && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && c.A3Nome == vende
                                 select new RelatorioFaturamentoDental
                                 {
                                     D2_FILIAL = SD20.D2Filial,
                                     D2_CLIENTE = SD20.D2Cliente,
                                     D2_LOJA = SD20.D2Emissao,
                                     A1_NOME = SA10.A1Nome,
                                     A1_CLINTER = SA10.A1Clinter,
                                     D2_DOC = SD20.D2Doc,
                                     D2_SERIE = SD20.D2Serie,
                                     D2_EMISSAO = SD20.D2Emissao,
                                     D2_PEDIDO = SD20.D2Pedido,
                                     C5_UNUMAGE = SC50.C5Unumage,
                                     C5_EMISSAO = SC50.C5Emissao,
                                     C5_VEND1 = SC50.C5Vend1,
                                     A3_NOME = c.A3Nome,
                                     C5_X_DTCIR = SC50.C5XDtcir,
                                     C5_X_NMMED = SC50.C5XNmmed,
                                     C5_X_NMPAC = SC50.C5XNmpac,
                                     C5_X_NMPLA = SC50.C5XNmpla,
                                     C5_UTPOPER = SC50.C5Utpoper,
                                     C6_PRODUTO = SC60.C6Produto,
                                     B1_DESC = SB10.B1Desc,
                                     D2_TOTAL = SD20.D2Total,
                                     D2_QUANT = SD20.D2Quant,
                                     D2_VALIPI = SD20.D2Valipi,
                                     D2_VALICM = SD20.D2Valicm,
                                     D2_DESCON = SD20.D2Descon
                                 }
                                  ).GroupBy(Prod =>
                                  new
                                  {
                                      Prod.D2_FILIAL,
                                      Prod.D2_CLIENTE,
                                      Prod.D2_LOJA,
                                      Prod.A1_NOME,
                                      Prod.A1_CLINTER,
                                      Prod.D2_DOC,
                                      Prod.D2_SERIE,
                                      Prod.D2_EMISSAO,
                                      Prod.D2_PEDIDO,
                                      Prod.C5_UNUMAGE,
                                      Prod.C5_EMISSAO,
                                      Prod.C5_VEND1,
                                      Prod.A3_NOME,
                                      Prod.C5_X_DTCIR,
                                      Prod.C5_X_NMMED,
                                      Prod.C5_X_NMPAC,
                                      Prod.C5_X_NMPLA,
                                      Prod.C5_UTPOPER,
                                      Prod.C6_PRODUTO,
                                      Prod.B1_DESC
                                  });

                    Relatorio = query.Select(x => new RelatorioFaturamentoDental
                    {
                        D2_FILIAL = x.Key.D2_FILIAL,
                        D2_CLIENTE = x.Key.D2_CLIENTE,
                        D2_LOJA = x.Key.D2_LOJA,
                        A1_NOME = x.Key.A1_NOME,
                        A1_CLINTER = x.Key.A1_CLINTER,
                        D2_DOC = x.Key.D2_DOC,
                        D2_SERIE = x.Key.D2_SERIE,
                        D2_EMISSAO = $"{x.Key.D2_EMISSAO.Substring(6, 2)}/{x.Key.D2_EMISSAO.Substring(4, 2)}/{x.Key.D2_EMISSAO.Substring(0, 4)}",
                        D2_PEDIDO = x.Key.D2_PEDIDO,
                        C5_UNUMAGE = x.Key.C5_UNUMAGE,
                        C5_EMISSAO = x.Key.C5_EMISSAO,
                        C5_VEND1 = x.Key.C5_VEND1,
                        A3_NOME = x.Key.A3_NOME,
                        C5_X_DTCIR = x.Key.C5_X_DTCIR,
                        C5_X_NMMED = x.Key.C5_X_NMMED,
                        C5_X_NMPAC = x.Key.C5_X_NMPAC,
                        C5_X_NMPLA = x.Key.C5_X_NMPLA,
                        C5_UTPOPER = x.Key.C5_UTPOPER,
                        C6_PRODUTO = x.Key.C6_PRODUTO,
                        B1_DESC = x.Key.B1_DESC,
                        D2_TOTAL = x.Sum(c => c.D2_TOTAL),
                        D2_QUANT = x.Sum(c => c.D2_QUANT),
                        D2_VALIPI = x.Sum(c => c.D2_VALIPI),
                        D2_VALICM = x.Sum(c => c.D2_VALICM),
                        D2_DESCON = x.Sum(c => c.D2_DESCON),
                    }).OrderBy(x => x.A3_NOME).ToList();


                    var CfDevolucao = new string[] { "1202", "2202", "3202", "1553", "2553" };

                    var queryDevolucao = (from SD10 in Protheus.Sd1010s
                                          join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Fornece = SF20.F2Cliente, Loja = SF20.F2Loja }
                                          join SD20 in Protheus.Sd2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Fornece = SD10.D1Fornece, Loja = SD10.D1Loja, Item = SD10.D1Itemori } equals new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Fornece = SD20.D2Cliente, Loja = SD20.D2Loja, Item = SD20.D2Item }
                                          join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                                          join SA10 in Protheus.Sa1010s on new { Fornece = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Fornece = SA10.A1Cod, Loja = SA10.A1Loja }
                                          join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                          where SD10.DELET != "*" && SF20.DELET != "*" && SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SC50.DELET != "*"
                                          && CfDevolucao.Contains(SD10.D1Cf) && SC50.C5Xtipopv == "D"
                                          && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                                          && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                                          && SC50.C5Nomvend == vende
                                          select new
                                          {
                                              SD10.D1Filial,
                                              SD10.D1Fornece,
                                              SD10.D1Loja,
                                              SA10.A1Nome,
                                              SA10.A1Clinter,
                                              SD10.D1Doc,
                                              SD10.D1Serie,
                                              SD10.D1Emissao,
                                              SD20.D2Doc,
                                              SD20.D2Emissao,
                                              SD10.D1Total,
                                              SD10.D1Valipi,
                                              SD10.D1Valicm,
                                              SC50.C5Vend1,
                                              SC50.C5Nomvend,
                                              SD10.D1Valdesc,
                                              SD10.D1Dtdigit,
                                              SA10.A1Est,
                                              SA10.A1Mun
                                          })
                                          .GroupBy(x => new
                                          {
                                              x.D1Filial,
                                              x.D1Fornece,
                                              x.D1Loja,
                                              x.A1Nome,
                                              x.A1Clinter,
                                              x.A1Est,
                                              x.A1Mun,
                                              x.D1Doc,
                                              x.D1Serie,
                                              x.D1Emissao,
                                              x.D1Dtdigit,
                                              x.D2Doc,
                                              x.D2Emissao,
                                              x.C5Vend1,
                                              x.C5Nomvend
                                          });

                    Devolucao = queryDevolucao.Select(x => new RelatorioDevolucaoDental
                    {
                        Emissao = $"{x.Key.D1Emissao.Substring(6, 2)}/{x.Key.D1Emissao.Substring(4, 2)}/{x.Key.D1Emissao.Substring(0, 4)}",
                        Filial = x.Key.D1Filial,
                        NumNF = x.Key.D1Doc,
                        EmissaoOrig = $"{x.Key.D2Emissao.Substring(6, 2)}/{x.Key.D2Emissao.Substring(4, 2)}/{x.Key.D2Emissao.Substring(0, 4)}",
                        NFOrig = x.Key.D2Doc,
                        CodVend = x.Key.C5Vend1,
                        Vendedor = x.Key.C5Nomvend,
                        Cliente = x.Key.A1Nome,
                        Total = -(x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valipi)),
                    }).OrderBy(x => x.Vendedor).ToList();
                }


                var vendedores = Relatorio.Select(Pedido => new RelatorioVendedorDental
                {
                    Vendedor = Pedido.A3_NOME,
                    Faturado = 0.00,
                    Devolucao = 0.00,
                    Comissao = 0.00,
                    Login = Pedido.A3_LOGIN
                }).DistinctBy(x => x.Vendedor).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Faturamento Dental");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Faturamento Dental");

                double Faturado = 0;
                double Devoluca = 0;

                sheet.Cells[2, 1].Value = "Faturado";
                sheet.Cells[2, 2].Value = "Devolução";
                sheet.Cells[2, 3].Value = "Total";

                int i = 5;

                sheet.Cells[i, 1].Value = "Faturado";
                sheet.Cells[i, 1, i, 9].Merge = true;

                i++;

                sheet.Cells[i, 1].Value = "Emissao NF";
                sheet.Cells[i, 2].Value = "Filial";
                sheet.Cells[i, 3].Value = "Num.NF";
                sheet.Cells[i, 4].Value = "Num.Pedido";
                sheet.Cells[i, 5].Value = "Vendedor";
                sheet.Cells[i, 6].Value = "Cliente";
                sheet.Cells[i, 7].Value = "Tipo de Operação";
                sheet.Cells[i, 8].Value = "Produto";
                sheet.Cells[i, 9].Value = "Desc.Produto";
                sheet.Cells[i, 10].Value = "QTD Faturada";
                sheet.Cells[i, 11].Value = "Total NF";
                sheet.Cells[i, 12].Value = "Comissão";

                i++;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.D2_EMISSAO;
                    sheet.Cells[i, 2].Value = Pedido.D2_FILIAL;
                    sheet.Cells[i, 3].Value = Pedido.D2_DOC;
                    sheet.Cells[i, 4].Value = Pedido.D2_PEDIDO;
                    sheet.Cells[i, 5].Value = Pedido.A3_NOME;
                    sheet.Cells[i, 6].Value = Pedido.A1_NOME;
                    sheet.Cells[i, 7].Value = Pedido.C5_UTPOPER;
                    sheet.Cells[i, 8].Value = Pedido.C6_PRODUTO;
                    sheet.Cells[i, 9].Value = Pedido.B1_DESC;

                    sheet.Cells[i, 10].Value = Pedido.D2_QUANT;
                    sheet.Cells[i, 10].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 11].Value = Pedido.D2_TOTAL;
                    sheet.Cells[i, 11].Style.Numberformat.Format = "0.00";

                    switch (Pedido.A3_LOGIN) {

                        case "PEDRO.KOVACS":
                        case "WILSON.GORGONE":
                        case "ROGER.MACEDO":
                        case "CIRO.PALLU":
                        {

                           sheet.Cells[i, 12].Value = Pedido.D2_TOTAL * 0.10;
                           sheet.Cells[i, 12].Style.Numberformat.Format = "0.00";
                           break;
                        }
                        case "KELLY.GONCALVES":
                        case "ROSANA.LUZ":
                        {
                            sheet.Cells[i, 12].Value = Pedido.D2_TOTAL * 0.02;
                            sheet.Cells[i, 12].Style.Numberformat.Format = "0.00";
                            break;
                        }
                        default:
                        {
                            sheet.Cells[i, 12].Value = Pedido.D2_TOTAL * 0.03;
                            sheet.Cells[i, 12].Style.Numberformat.Format = "0.00";
                            break;
                        }
                    }

                    Faturado += Pedido.D2_TOTAL;

                    var teste = vendedores.FirstOrDefault(x => x.Vendedor == Pedido.A3_NOME);

                    if (teste != null)
                    {
                        teste.Faturado += Pedido.D2_TOTAL;
                    }

                    i++;
                });
                i++;

                sheet.Cells[i, 1].Value = "Devolução";
                sheet.Cells[i, 1, i, 9].Merge = true;

                i++;

                sheet.Cells[i, 1].Value = "Emissão";
                sheet.Cells[i, 2].Value = "Filial";
                sheet.Cells[i, 3].Value = "Num. NF";
                sheet.Cells[i, 4].Value = "Emissão Orig.";
                sheet.Cells[i, 5].Value = "NF. Orig.";
                sheet.Cells[i, 6].Value = "Cod Vend";
                sheet.Cells[i, 7].Value = "Vendedor";
                sheet.Cells[i, 8].Value = "Cliente";
                sheet.Cells[i, 9].Value = "Total";
                sheet.Cells[i, 10].Value = "Comissão";
                i++;

                Devolucao.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Emissao;
                    sheet.Cells[i, 2].Value = Pedido.Filial;
                    sheet.Cells[i, 3].Value = Pedido.NumNF;
                    sheet.Cells[i, 4].Value = Pedido.EmissaoOrig;
                    sheet.Cells[i, 5].Value = Pedido.NFOrig;
                    sheet.Cells[i, 6].Value = Pedido.CodVend;
                    sheet.Cells[i, 7].Value = Pedido.Vendedor;
                    sheet.Cells[i, 8].Value = Pedido.Cliente;

                    sheet.Cells[i, 9].Value = Pedido.Total;
                    sheet.Cells[i, 9].Style.Numberformat.Format = "0.00";

                    var teste = vendedores.FirstOrDefault(x => x.Vendedor == Pedido.Vendedor);

                    if (teste != null)
                    {
                        teste.Devolucao += Pedido.Total;

                        switch (teste.Login)
                        {

                            case "PEDRO.KOVACS":
                            case "WILSON.GORGONE":
                            case "ROGER.MACEDO":
                            case "CIRO.PALLU":
                                {

                                    sheet.Cells[i, 12].Value = Pedido.Total * 0.10;
                                    sheet.Cells[i, 12].Style.Numberformat.Format = "0.00";
                                    break;
                                }
                            case "KELLY.GONCALVES":
                            case "ROSANA.LUZ":
                                {
                                    sheet.Cells[i, 12].Value = Pedido.Total * 0.02;
                                    sheet.Cells[i, 12].Style.Numberformat.Format = "0.00";
                                    break;
                                }
                            default:
                                {
                                    sheet.Cells[i, 12].Value = Pedido.Total * 0.03;
                                    sheet.Cells[i, 12].Style.Numberformat.Format = "0.00";
                                    break;
                                }
                        }
                    }

                    Devoluca += Pedido.Total;

                    i++;
                });

                sheet.Cells[3, 1].Value = Faturado;
                sheet.Cells[3, 1].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                sheet.Cells[3, 2].Value = Devoluca;
                sheet.Cells[3, 2].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                sheet.Cells[3, 3].Value = Faturado+Devoluca;
                sheet.Cells[3, 3].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                i++;

                sheet.Cells[i, 1].Value = "Comissões";
                sheet.Cells[i, 1, i, 9].Merge = true;

                i++;

                sheet.Cells[i, 1].Value = "Vendedor";
                sheet.Cells[i, 2].Value = "Faturado";
                sheet.Cells[i, 3].Value = "Devolução";
                sheet.Cells[i, 4].Value = "Comissao";

                i++;

                vendedores.ForEach(Vendedor =>
                {
                    sheet.Cells[i, 1].Value = Vendedor.Vendedor;

                    sheet.Cells[i, 2].Value =  Vendedor.Faturado;
                    sheet.Cells[i, 2].Style.Numberformat.Format = "0.00";

                    sheet.Cells[i, 3].Value = Vendedor.Devolucao;
                    sheet.Cells[i, 3].Style.Numberformat.Format = "0.00";

                    switch (Vendedor.Login)
                    {

                        case "PEDRO.KOVACS":
                        case "WILSON.GORGONE":
                        case "ROGER.MACEDO":
                        case "CIRO.PALLU":
                            {

                                sheet.Cells[i, 4].Value = (Vendedor.Faturado + Vendedor.Devolucao) * 0.10;
                                sheet.Cells[i, 4].Style.Numberformat.Format = "0.00";
                                break;
                            }
                        case "KELLY.GONCALVES":
                        case "ROSANA.LUZ":
                            {
                                sheet.Cells[i, 4].Value = (Vendedor.Faturado + Vendedor.Devolucao) * 0.02;
                                sheet.Cells[i, 4].Style.Numberformat.Format = "0.00";
                                break;
                            }
                        default:
                            {
                                sheet.Cells[i, 4].Value = (Vendedor.Faturado + Vendedor.Devolucao) * 0.03;
                                sheet.Cells[i, 4].Style.Numberformat.Format = "0.00";
                                break;
                            }
                    }

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FaturamentoDental.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoDental Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
