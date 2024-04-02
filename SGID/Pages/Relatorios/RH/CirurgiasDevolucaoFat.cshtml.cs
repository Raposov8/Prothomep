using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Models.Denuo;
using SGID.Models;
using SGID.Data.Models;

namespace SGID.Pages.Relatorios.RH
{
    [Authorize]
    public class CirurgiasDevolucaoFatModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioDevolucaoFat> Relatorio { get; set; } = new List<RelatorioDevolucaoFat>();

        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public CirurgiasDevolucaoFatModel(TOTVSDENUOContext context, ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }

        public void OnGet(int id)
        {

        }

        public IActionResult OnPostAsync(DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;
                Relatorio = new List<RelatorioDevolucaoFat>();

                var user = User.Identity.Name.Split("@")[0].ToUpper();

                if (User.IsInRole("GestorComercial"))
                {

                    if (user != "TIAGO.FONSECA")
                    {
                        #region Gestor


                        var teste = (from SD10 in Protheus.Sd1010s
                                     join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SA30 in Protheus.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                     join SC50 in Protheus.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                     join SA10 in Protheus.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                     where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
                                     && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                     && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                                     && SC50.C5Utpoper == "F"
                                     && (int)(object)SD10.D1Dtdigit >= 20200801 && (SA30.A3Xlogin == user || SA30.A3Xlogsup == user)
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
                                         SA30.A3Xdescun
                                     }).GroupBy(x => new
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
                                         x.A3Xdescun
                                     });

                        Relatorio = teste.Select(x => new RelatorioDevolucaoFat
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
                            Linha = x.Key.A3Xdescun
                        }).ToList();

                        #endregion
                    }
                    else
                    {
                        #region Tiago

                        var teste = (from SD10 in Protheus.Sd1010s
                                     join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SA30 in Protheus.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                     join SC50 in Protheus.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                     join SA10 in Protheus.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                     where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
                                     && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                     && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                                     && SC50.C5Utpoper == "F"
                                     && (int)(object)SD10.D1Dtdigit >= 20200801 && SA30.A3Xlogsup != "ANDRE.SALES"
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
                                         SA30.A3Xdescun
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
                                 x.A3Xdescun
                             });

                        Relatorio = teste.Select(x => new RelatorioDevolucaoFat
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
                            Linha = x.Key.A3Xdescun
                        }).ToList();

                        #endregion
                    }
                }
                else
                {
                    var teste = (from SD10 in Protheus.Sd1010s
                                 join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                 join SA30 in Protheus.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                 join SC50 in Protheus.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                 join SA10 in Protheus.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                 join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                 where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
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
                                     SA30.A3Xdescun
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
                                 x.A3Xdescun
                             });

                    Relatorio = teste.Select(x => new RelatorioDevolucaoFat
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
                        Linha = x.Key.A3Xdescun
                    }).ToList();
                }
                

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasDevolucaoFatVendedor", user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(DateTime DataInicio,DateTime DataFim)
        {
            try
            {
                Inicio = DataInicio;
                Fim = DataFim;
                Relatorio = new List<RelatorioDevolucaoFat>();

                var user = User.Identity.Name.Split("@")[0].ToUpper();

                if (User.IsInRole("GestorComercial"))
                {

                    if (user != "TIAGO.FONSECA")
                    {
                        #region Gestor


                        var teste = (from SD10 in Protheus.Sd1010s
                                     join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SA30 in Protheus.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                     join SC50 in Protheus.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                     join SA10 in Protheus.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                     where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
                                     && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                     && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                                     && SC50.C5Utpoper == "F"
                                     && (int)(object)SD10.D1Dtdigit >= 20200801 && (SA30.A3Xlogin == user || SA30.A3Xlogsup == user)
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
                                         SA30.A3Xdescun
                                     }).GroupBy(x => new
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
                                         x.A3Xdescun
                                     });

                        Relatorio = teste.Select(x => new RelatorioDevolucaoFat
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
                            Linha = x.Key.A3Xdescun
                        }).ToList();

                        #endregion
                    }
                    else
                    {
                        #region Tiago

                        var teste = (from SD10 in Protheus.Sd1010s
                                     join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SA30 in Protheus.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                     join SC50 in Protheus.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                     join SA10 in Protheus.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                     where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
                                     && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                                     && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                                     && SC50.C5Utpoper == "F"
                                     && (int)(object)SD10.D1Dtdigit >= 20200801 && SA30.A3Xlogsup != "ANDRE.SALES"
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
                                         SA30.A3Xdescun
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
                                 x.A3Xdescun
                             });

                        Relatorio = teste.Select(x => new RelatorioDevolucaoFat
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
                            Linha = x.Key.A3Xdescun
                        }).ToList();

                        #endregion
                    }
                }
                else
                {
                    var teste = (from SD10 in Protheus.Sd1010s
                                 join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                                 join SA30 in Protheus.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                                 join SC50 in Protheus.Sc5010s on new { Filial = SD10.D1Filial, Nota = SD10.D1Nfori } equals new { Filial = SC50.C5Filial, Nota = SC50.C5Nota }
                                 join SA10 in Protheus.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                                 join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                 where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SA30.DELET != "*"
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
                                     SA30.A3Xdescun
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

                             });

                    Relatorio = teste.Select(x => new RelatorioDevolucaoFat
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
                        Linha = x.Key.A3Xdescun
                    }).ToList();
                }

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Cirurgias Devolucao Fat");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Cirurgias Devolucao Fat");

                sheet.Cells[1, 1].Value = "FILIAL";
                sheet.Cells[1, 2].Value = "CLIFOR";
                sheet.Cells[1, 3].Value = "LOJA";
                sheet.Cells[1, 4].Value = "NOME";
                sheet.Cells[1, 5].Value = "TIPO";
                sheet.Cells[1, 6].Value = "NF";
                sheet.Cells[1, 7].Value = "SERIE";
                sheet.Cells[1, 8].Value = "DIGITACAO";
                sheet.Cells[1, 9].Value = "TOTAL";
                sheet.Cells[1, 10].Value = "VALIPI";
                sheet.Cells[1, 11].Value = "VALICM";
                sheet.Cells[1, 12].Value = "DESCON";
                sheet.Cells[1, 13].Value = "TOTALBRUT";
                sheet.Cells[1, 14].Value = "A3_NOME";
                sheet.Cells[1, 15].Value = "D1_NFORI";
                sheet.Cells[1, 16].Value = "D1_SERIORI";
                sheet.Cells[1, 17].Value = "D1_DATORI";

                int i = 2;

                Relatorio.ForEach(Pedido =>
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

                    i++;
                });


                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CirurgiasDevolucaoFatVendedor.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasDevolucaoFatVendedor Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
