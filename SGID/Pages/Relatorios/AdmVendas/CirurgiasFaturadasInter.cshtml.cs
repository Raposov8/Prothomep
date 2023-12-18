using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.AdmVendas
{
    public class CirurgiasFaturadasInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioCirurgiasFaturadas> Relatorio = new List<RelatorioCirurgiasFaturadas>();

        public List<RelatorioDevolucaoFat> Relatorio2 = new List<RelatorioDevolucaoFat>();


        public CirurgiasFaturadasInterModel(TOTVSINTERContext context, ApplicationDbContext sgid)
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

                int[] CF = new int[] { 5551, 6551, 6107, 6109 };
                var user = User.Identity.Name.Split("@")[0].ToUpper();

                //&& (SA10.A1Xgrinte != "000011" || SA10.A1Xgrinte != "000012")


                var query = (from SD20 in Protheus.Sd2010s
                             join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                             && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                             (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", ""))
                             && SD20.D2Quant != 0 && SC50.C5Utpoper == "F"  && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
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
                                 SA30.A3Nome,
                                 DataCirurgia = SC50.C5XDtcir,
                                 NomMed = SC50.C5XNmmed,
                                 NomPac = SC50.C5XNmpac,
                                 NomPla = SC50.C5XNmpla,
                                 SC50.C5Utpoper,
                                 SA30.A3Xdescun,
                                 SC50.C5Nomclie,
                                 SB10.B1Fabric
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
                    x.A3Nome,
                    x.DataCirurgia,
                    x.NomMed,
                    x.NomPac,
                    x.NomPla,
                    x.C5Utpoper,
                    x.A3Xdescun,
                    x.C5Nomclie,
                    x.B1Fabric
                })
                .Select(x => new RelatorioCirurgiasFaturadas
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
                    A3Nome = x.Key.A3Nome,
                    XDtcir = $"{x.Key.DataCirurgia.Substring(6, 2)}/{x.Key.DataCirurgia.Substring(4, 2)}/{x.Key.DataCirurgia.Substring(0, 4)}",
                    XNMMed = x.Key.NomMed,
                    XNMPac = x.Key.NomPac,
                    XNMPla = x.Key.NomPla,
                    Utpoper = x.Key.C5Utpoper,
                    Linha = x.Key.A3Xdescun,
                    Entrega = x.Key.C5Nomclie,
                    Fornecedor = x.Key.B1Fabric
                }).OrderBy(x => x.A3Nome).ToList();


                Relatorio2 = new List<RelatorioDevolucaoFat>();

                var teste = (from SD10 in Protheus.Sd1010s
                             join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SA30 in Protheus.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                             join SC50 in Protheus.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                             join SA10 in Protheus.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                             where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                             && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                             && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             && (int)(object)SD10.D1Dtdigit >= 20200801 && SC50.C5Utpoper == "F"
                             orderby SA30.A3Nome
                             select new
                             {
                                 SD10.D1Filial,
                                 SD10.D1Fornece,
                                 SD10.D1Loja,
                                 SA10.A1Nome,
                                 SA10.A1Clinter,
                                 SD10.D1Doc,
                                 SD10.D1Serie,
                                 SD10.D1Dtdigit,
                                 SD10.D1Total,
                                 SD10.D1Valdesc,
                                 SD10.D1Valipi,
                                 SD10.D1Valicm,
                                 SA30.A3Nome,
                                 SD10.D1Nfori,
                                 SD10.D1Seriori,
                                 SD10.D1Datori,
                                 SD10.D1Emissao,
                                 SA30.A3Xdescun,
                                 SC50.C5Nomclie,
                                 SB10.B1Fabric
                             }
                         ).GroupBy(x => new
                         {
                             x.A1Nome,
                             x.A3Nome,
                             x.D1Filial,
                             x.D1Fornece,
                             x.D1Loja,
                             x.A1Clinter,
                             x.D1Doc,
                             x.D1Serie,
                             x.D1Emissao,
                             x.D1Dtdigit,
                             x.D1Nfori,
                             x.D1Seriori,
                             x.D1Datori,
                             x.A3Xdescun,
                             x.C5Nomclie,
                             x.B1Fabric
                         });

                Relatorio2 = teste.Select(x => new RelatorioDevolucaoFat
                {
                    Filial = x.Key.D1Filial,
                    Clifor = x.Key.D1Fornece,
                    Loja = x.Key.D1Loja,
                    Nome = x.Key.A1Nome,
                    Tipo = x.Key.A1Clinter,
                    Nf = x.Key.D1Doc,
                    Serie = x.Key.D1Serie,
                    Digitacao = x.Key.D1Dtdigit,
                    Total = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc),
                    Valipi = x.Sum(c => c.D1Valipi),
                    Valicm = x.Sum(c => c.D1Valicm),
                    Descon = x.Sum(c => c.D1Valdesc),
                    TotalBrut = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valdesc),
                    A3Nome = x.Key.A3Nome,
                    D1Nfori = x.Key.D1Nfori,
                    D1Seriori = x.Key.D1Seriori,
                    D1Datori = x.Key.D1Datori,
                    Linha = x.Key.A3Xdescun,
                    Entrega = x.Key.C5Nomclie,
                    Fornecedor = x.Key.B1Fabric
                }).ToList();


                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirugiasFaturadasInterADM", user);
                return LocalRedirect("/error");
            }
        }
        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Relatorio = new List<RelatorioCirurgiasFaturadas>();
                int[] CF = new int[] { 5551, 6551, 6107, 6109 };
                var user = User.Identity.Name.Split("@")[0].ToUpper();

                #region Faturamento
                var query = (from SD20 in Protheus.Sd2010s
                             join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SA30 in Protheus.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                             && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                             (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", ""))
                             && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
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
                                 SA30.A3Nome,
                                 DataCirurgia = SC50.C5XDtcir,
                                 NomMed = SC50.C5XNmmed,
                                 NomPac = SC50.C5XNmpac,
                                 NomPla = SC50.C5XNmpla,
                                 SC50.C5Utpoper,
                                 SA30.A3Xdescun,
                                 Entrega = SC50.C5Nomclie,
                                 Fornece = SB10.B1Fabric
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
                    x.A3Nome,
                    x.DataCirurgia,
                    x.NomMed,
                    x.NomPac,
                    x.NomPla,
                    x.C5Utpoper,
                    x.A3Xdescun,
                    x.Entrega,
                    x.Fornece
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
                    A3Nome = x.Key.A3Nome,
                    XDtcir = $"{x.Key.DataCirurgia.Substring(6, 2)}/{x.Key.DataCirurgia.Substring(4, 2)}/{x.Key.DataCirurgia.Substring(0, 4)}",
                    XNMMed = x.Key.NomMed,
                    XNMPac = x.Key.NomPac,
                    XNMPla = x.Key.NomPla,
                    Utpoper = x.Key.C5Utpoper,
                    Linha = x.Key.A3Xdescun,
                    Entrega = x.Key.Entrega,
                    Fornecedor = x.Key.Fornece
                }).OrderBy(x => x.A3Nome).ToList();
                #endregion

                #region Devolucao
                Relatorio2 = new List<RelatorioDevolucaoFat>();

                var teste = (from SD10 in Protheus.Sd1010s
                             join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                             join SA30 in Protheus.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                             join SC50 in Protheus.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                             join SA10 in Protheus.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                             where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                             && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                             && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             && (int)(object)SD10.D1Dtdigit >= 20200801 && SC50.C5Utpoper == "F"
                             orderby SA30.A3Nome
                             select new
                             {
                                 SD10.D1Filial,
                                 SD10.D1Fornece,
                                 SD10.D1Loja,
                                 SA10.A1Nome,
                                 SA10.A1Clinter,
                                 SD10.D1Doc,
                                 SD10.D1Serie,
                                 SD10.D1Dtdigit,
                                 SD10.D1Total,
                                 SD10.D1Valdesc,
                                 SD10.D1Valipi,
                                 SD10.D1Valicm,
                                 SA30.A3Nome,
                                 SD10.D1Nfori,
                                 SD10.D1Seriori,
                                 SD10.D1Datori,
                                 SD10.D1Emissao,
                                 SA30.A3Xdescun,
                                 SC50.C5Nomclie,
                                 SB10.B1Fabric
                             }
                         )
                         .GroupBy(x => new
                         {
                             x.A1Nome,
                             x.A3Nome,
                             x.D1Filial,
                             x.D1Fornece,
                             x.D1Loja,
                             x.A1Clinter,
                             x.D1Doc,
                             x.D1Serie,
                             x.D1Emissao,
                             x.D1Dtdigit,
                             x.D1Nfori,
                             x.D1Seriori,
                             x.D1Datori,
                             x.A3Xdescun,
                             x.C5Nomclie,
                             x.B1Fabric
                         });

                Relatorio2 = teste.Select(x => new RelatorioDevolucaoFat
                {
                    Filial = x.Key.D1Filial,
                    Clifor = x.Key.D1Fornece,
                    Loja = x.Key.D1Loja,
                    Nome = x.Key.A1Nome,
                    Tipo = x.Key.A1Clinter,
                    Nf = x.Key.D1Doc,
                    Serie = x.Key.D1Serie,
                    Digitacao = x.Key.D1Dtdigit,
                    Total = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc),
                    Valipi = x.Sum(c => c.D1Valipi),
                    Valicm = x.Sum(c => c.D1Valicm),
                    Descon = x.Sum(c => c.D1Valdesc),
                    TotalBrut = x.Sum(c => c.D1Total) - x.Sum(c => c.D1Valdesc) + x.Sum(c => c.D1Valdesc),
                    A3Nome = x.Key.A3Nome,
                    D1Nfori = x.Key.D1Nfori,
                    D1Seriori = x.Key.D1Seriori,
                    D1Datori = x.Key.D1Datori,
                    Linha = x.Key.A3Xdescun,
                    Entrega = x.Key.C5Nomclie,
                    Fornecedor = x.Key.B1Fabric
                }).ToList();


                #endregion

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Cirurgias Faturadas");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Cirurgias Faturadas");

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
                sheet.Cells[1, 17].Value = "Especialidade";
                sheet.Cells[1, 18].Value = "Cliente Entrega";
                sheet.Cells[1, 19].Value = "Fornecedor";

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
                    sheet.Cells[i, 17].Value = Pedido.Linha;
                    sheet.Cells[i, 18].Value = Pedido.Entrega;
                    sheet.Cells[i, 19].Value = Pedido.Fornecedor;

                    i++;
                });


                var format = sheet.Cells[i, 8, i, 9];
                format.Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                i++;
                i++;

                sheet.Cells[i, 1].Value = "FILIAL";
                sheet.Cells[i, 2].Value = "CLIFOR";
                sheet.Cells[i, 3].Value = "LOJA";
                sheet.Cells[i, 4].Value = "NOME";
                sheet.Cells[i, 5].Value = "TIPO";
                sheet.Cells[i, 6].Value = "NF";
                sheet.Cells[i, 7].Value = "SERIE";
                sheet.Cells[i, 8].Value = "DIGITACAO";
                sheet.Cells[i, 9].Value = "TOTAL";
                sheet.Cells[i, 10].Value = "VALIPI";
                sheet.Cells[i, 11].Value = "VALICM";
                sheet.Cells[i, 12].Value = "DESCON";
                sheet.Cells[i, 13].Value = "TOTALBRUT";
                sheet.Cells[i, 14].Value = "A3_NOME";
                sheet.Cells[i, 15].Value = "D1_NFORI";
                sheet.Cells[i, 16].Value = "D1_SERIORI";
                sheet.Cells[i, 17].Value = "D1_DATORI";
                sheet.Cells[i, 18].Value = "Especialidade";
                sheet.Cells[i, 19].Value = "Cliente Entrega";
                sheet.Cells[i, 20].Value = "Fornecedor";

                i++;

                Relatorio2.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.Clifor;
                    sheet.Cells[i, 3].Value = Pedido.Loja;
                    sheet.Cells[i, 4].Value = Pedido.Nome;
                    sheet.Cells[i, 5].Value = Pedido.Tipo;
                    sheet.Cells[i, 6].Value = Pedido.Nf;
                    sheet.Cells[i, 7].Value = Pedido.Serie;
                    sheet.Cells[i, 8].Value = Pedido.Digitacao;
                    sheet.Cells[i, 9].Value = Pedido.Total;
                    sheet.Cells[i, 10].Value = Pedido.Valipi;
                    sheet.Cells[i, 11].Value = Pedido.Valicm;
                    sheet.Cells[i, 12].Value = Pedido.Descon;
                    sheet.Cells[i, 13].Value = Pedido.TotalBrut;
                    sheet.Cells[i, 14].Value = Pedido.A3Nome;
                    sheet.Cells[i, 15].Value = Pedido.D1Nfori;
                    sheet.Cells[i, 16].Value = Pedido.D1Seriori;
                    sheet.Cells[i, 17].Value = Pedido.D1Datori;
                    sheet.Cells[i, 18].Value = Pedido.Linha;
                    sheet.Cells[i, 19].Value = Pedido.Entrega;
                    sheet.Cells[i, 20].Value = Pedido.Fornecedor;


                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CirurgiasFaturadasInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasFaturadasInterADM Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
