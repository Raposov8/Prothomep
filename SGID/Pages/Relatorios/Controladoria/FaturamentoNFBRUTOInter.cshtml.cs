using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Controladoria.FaturamentoNF;
using SGID.Models.Controladoria.FaturamentoNF.RegistrosFaturamento;
using SGID.Models.Inter;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class FaturamentoNFBRUTOInterModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext Protheus { get; set; }

        public List<RelatorioNF> Relatorio { get; set; } = new List<RelatorioNF>();
        public List<RelatorioFaturamentoNF> Faturamento { get; set; } = new List<RelatorioFaturamentoNF>();
        public List<RelatorioDevolucaoNF> Devolucao { get; set; } = new List<RelatorioDevolucaoNF>();

        public string Ano { get; set; }

        public FaturamentoNFBRUTOInterModel(ApplicationDbContext sgid,TOTVSINTERContext inter)
        {
            SGID = sgid;
            Protheus = inter;
        }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string Ano)
        {
            try
            {
                this.Ano = Ano;
                var InicioAno = Convert.ToInt32($"{Ano}0101");
                var FimAno = Convert.ToInt32($"{Ano}1231");


                var CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };
                var CfNe = new int[] { 1202, 1553, 2202, 2553 };

                #region Faturado

                var queryFaturado = (from SD20 in Protheus.Sd2010s
                                     join SA10 in Protheus.Sa1010s on new { Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                     join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SC50 in Protheus.Sc5010s on new { Pedido = SD20.D2Pedido, Filial = SD20.D2Filial, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Pedido = SC50.C5Num, Filial = SC50.C5Filial, Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                                     where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" &&
                                     (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114)
                                     || ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf))
                                     && (int)(object)SD20.D2Emissao >= InicioAno && (int)(object)SD20.D2Emissao <= FimAno && SD20.D2Quant != 0
                                     select new
                                     {
                                         Filial = SD20.D2Filial,
                                         Cliente = SD20.D2Cliente,
                                         Loja = SD20.D2Loja,
                                         Nome = SA10.A1Nome,
                                         Tipo = SA10.A1Clinter,
                                         Est = SA10.A1Est,
                                         Mun = SA10.A1Mun,
                                         NF = SD20.D2Doc,
                                         Serie = SD20.D2Serie,
                                         Emissao = SD20.D2Emissao,
                                         Total = SD20.D2Total,
                                         Valipi = SD20.D2Valipi,
                                         Valicm = SD20.D2Valicm,
                                         Descon = SD20.D2Descon,
                                         TotalBrut = SD20.D2Valbrut
                                     })
                                            .GroupBy(x => new
                                            {
                                                x.Filial,
                                                x.Cliente,
                                                x.Loja,
                                                x.Nome,
                                                x.Tipo,
                                                x.Est,
                                                x.Mun,
                                                x.NF,
                                                x.Serie,
                                                x.Emissao
                                            })
                                            .Select(x => new RegistrosNF
                                            {
                                                Filial = x.Key.Filial,
                                                CliFor = x.Key.Cliente,
                                                Loja = x.Key.Loja,
                                                Nome = x.Key.Nome,
                                                Tipo = x.Key.Tipo == "H" ? "HOSPITAL" : x.Key.Tipo == "M" ? "MEDICO" : x.Key.Tipo == "I" ? "INSTRUMENTADOR" : x.Key.Tipo == "N" ? "NORMAL" : x.Key.Tipo == "C" ? "CONVENIO" : x.Key.Tipo == "P" ? "PARTICULAR" : x.Key.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                                NF = x.Key.NF,
                                                Serie = x.Key.Serie,
                                                Emissao = $"{x.Key.Emissao.Substring(6, 2)}/{x.Key.Emissao.Substring(4, 2)}/{x.Key.Emissao.Substring(0, 4)}",
                                                Total = x.Sum(c => c.TotalBrut),
                                                Valipi = x.Sum(c => c.Valipi),
                                                Valicm = x.Sum(c => c.Valicm),
                                                Descon = x.Sum(c => c.Descon),
                                                Mes = x.Key.Emissao.Substring(4, 2)
                                            }).ToList();

                var TipoFaturados = queryFaturado.Select(x => x.Tipo).Distinct().ToList();
                queryFaturado.ForEach(c =>
                {

                    if (!Faturamento.Any(x => x.Tipo == c.Tipo))
                    {
                        var Todos = queryFaturado.Where(x => x.Tipo == c.Tipo).ToList();

                        var NovoRelatorio = new RelatorioFaturamentoNF { Tipo = c.Tipo };

                        Todos.ForEach(d =>
                        {
                            switch (d.Mes)
                            {
                                case "01": NovoRelatorio.Janeiro += d.Total.Value; break;
                                case "02": NovoRelatorio.Fevereiro += d.Total.Value; break;
                                case "03": NovoRelatorio.Marco += d.Total.Value; break;
                                case "04": NovoRelatorio.Abril += d.Total.Value; break;
                                case "05": NovoRelatorio.Maio += d.Total.Value; break;
                                case "06": NovoRelatorio.Junho += d.Total.Value; break;
                                case "07": NovoRelatorio.Julho += d.Total.Value; break;
                                case "08": NovoRelatorio.Agosto += d.Total.Value; break;
                                case "09": NovoRelatorio.Setembro += d.Total.Value; break;
                                case "10": NovoRelatorio.Outubro += d.Total.Value; break;
                                case "11": NovoRelatorio.Novembro += d.Total.Value; break;
                                case "12": NovoRelatorio.Dezembro += d.Total.Value; break;
                            }

                            NovoRelatorio.Total += d.Total.Value;
                        });

                        Faturamento.Add(NovoRelatorio);
                    }
                });
                Faturamento = Faturamento.OrderBy(x => x.Tipo).ToList();
                #endregion

                #region Devolucao
                var queryDevolucao = (from SD10 in Protheus.Sd1010s
                                      join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, NF = SD10.D1Nfori, Serie = SD10.D1Seriori, Forne = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, NF = SF20.F2Doc, Serie = SF20.F2Serie, Forne = SF20.F2Cliente, Loja = SF20.F2Loja } into Sr
                                      from A in Sr.DefaultIfEmpty()
                                      join SA10 in Protheus.Sa1010s on new { Forne = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forne = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                      where SD10.DELET != "*" && A.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*"
                                      && CfNe.Contains((int)(object)SD10.D1Cf) && (int)(object)SD10.D1Dtdigit >= InicioAno && (int)(object)SD10.D1Dtdigit <= FimAno
                                      select new
                                      {
                                          Filial = SD10.D1Filial,
                                          Cliente = SD10.D1Fornece,
                                          Loja = SD10.D1Loja,
                                          Nome = SA10.A1Nome,
                                          Tipo = SA10.A1Clinter,
                                          Est = SA10.A1Est,
                                          Mun = SA10.A1Mun,
                                          NF = SD10.D1Doc,
                                          Serie = SD10.D1Serie,
                                          Emissao = SD10.D1Emissao,
                                          Total = SD10.D1Total - SD10.D1Valdesc,
                                          Valipi = SD10.D1Valipi,
                                          Valicm = SD10.D1Valicm,
                                          Descon = SD10.D1Valdesc,
                                          DTDIGIT = SD10.D1Dtdigit,
                                          TotalBruto = SD10.D1Total - SD10.D1Valdesc + SD10.D1Valipi
                                      }
                                         ).GroupBy(x => new
                                         {
                                             x.Filial,
                                             x.Cliente,
                                             x.Loja,
                                             x.Nome,
                                             x.Tipo,
                                             x.Est,
                                             x.Mun,
                                             x.NF,
                                             x.Serie,
                                             x.Emissao,
                                             x.DTDIGIT
                                         }).Select(x => new RegistrosNF
                                         {

                                             Filial = x.Key.Filial,
                                             CliFor = x.Key.Cliente,
                                             Loja = x.Key.Loja,
                                             Nome = x.Key.Nome,
                                             Tipo = x.Key.Tipo == "H" ? "HOSPITAL" : x.Key.Tipo == "M" ? "MEDICO" : x.Key.Tipo == "I" ? "INSTRUMENTADOR" : x.Key.Tipo == "N" ? "NORMAL" : x.Key.Tipo == "C" ? "CONVENIO" : x.Key.Tipo == "P" ? "PARTICULAR" : x.Key.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                             NF = x.Key.NF,
                                             Serie = x.Key.Serie,
                                             Emissao = $"{x.Key.Emissao.Substring(6, 2)}/{x.Key.Emissao.Substring(4, 2)}/{x.Key.Emissao.Substring(0, 4)}",
                                             Total = -x.Sum(c => c.TotalBruto),
                                             Valipi = -x.Sum(c => c.Valipi),
                                             Valicm = -x.Sum(c => c.Valicm),
                                             Descon = -x.Sum(c => c.Descon),
                                             Mes = x.Key.DTDIGIT.Substring(4, 2)
                                         }).ToList();


                queryDevolucao.ForEach(c =>
                {

                    if (!Devolucao.Any(x => x.Tipo == c.Tipo))
                    {
                        var Todos = queryDevolucao.Where(x => x.Tipo == c.Tipo).ToList();

                        var NovoRelatorio = new RelatorioDevolucaoNF { Tipo = c.Tipo };

                        Todos.ForEach(d =>
                        {
                            switch (d.Mes)
                            {
                                case "01": NovoRelatorio.Janeiro += d.Total.Value; break;
                                case "02": NovoRelatorio.Fevereiro += d.Total.Value; break;
                                case "03": NovoRelatorio.Marco += d.Total.Value; break;
                                case "04": NovoRelatorio.Abril += d.Total.Value; break;
                                case "05": NovoRelatorio.Maio += d.Total.Value; break;
                                case "06": NovoRelatorio.Junho += d.Total.Value; break;
                                case "07": NovoRelatorio.Julho += d.Total.Value; break;
                                case "08": NovoRelatorio.Agosto += d.Total.Value; break;
                                case "09": NovoRelatorio.Setembro += d.Total.Value; break;
                                case "10": NovoRelatorio.Outubro += d.Total.Value; break;
                                case "11": NovoRelatorio.Novembro += d.Total.Value; break;
                                case "12": NovoRelatorio.Dezembro += d.Total.Value; break;
                            }

                            NovoRelatorio.Total += d.Total.Value;
                        });

                        Devolucao.Add(NovoRelatorio);
                    }
                });

                Devolucao = Devolucao.OrderBy(x => x.Tipo).ToList();
                #endregion

                #region Relatorio
                var query = queryFaturado.Concat(queryDevolucao).ToList();

                var Tipos = query.Select(x => x.Tipo).Distinct().ToList();

                query.ForEach(c =>
                {

                    if (!Relatorio.Any(x => x.Tipo == c.Tipo))
                    {
                        var Todos = query.Where(x => x.Tipo == c.Tipo).ToList();

                        var NovoRelatorio = new RelatorioNF { Tipo = c.Tipo };

                        Todos.ForEach(d =>
                        {
                            switch (d.Mes)
                            {
                                case "01": NovoRelatorio.Janeiro += d.Total.Value; break;
                                case "02": NovoRelatorio.Fevereiro += d.Total.Value; break;
                                case "03": NovoRelatorio.Marco += d.Total.Value; break;
                                case "04": NovoRelatorio.Abril += d.Total.Value; break;
                                case "05": NovoRelatorio.Maio += d.Total.Value; break;
                                case "06": NovoRelatorio.Junho += d.Total.Value; break;
                                case "07": NovoRelatorio.Julho += d.Total.Value; break;
                                case "08": NovoRelatorio.Agosto += d.Total.Value; break;
                                case "09": NovoRelatorio.Setembro += d.Total.Value; break;
                                case "10": NovoRelatorio.Outubro += d.Total.Value; break;
                                case "11": NovoRelatorio.Novembro += d.Total.Value; break;
                                case "12": NovoRelatorio.Dezembro += d.Total.Value; break;
                            }

                            NovoRelatorio.Total += d.Total.Value;
                        });

                        Relatorio.Add(NovoRelatorio);
                    }

                });

                Relatorio = Relatorio.OrderBy(x => x.Tipo).ToList();
                #endregion


                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoNFBRUTOInter",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(string Ano)
        {
            try
            {

                var InicioAno = Convert.ToInt32($"{Ano}0101");
                var FimAno = Convert.ToInt32($"{Ano}1231");


                var CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };
                var CfNe = new int[] { 1202, 1553, 2202, 2553 };

                #region Faturado

                var queryFaturado = (from SD20 in Protheus.Sd2010s
                                     join SA10 in Protheus.Sa1010s on new { Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                     join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SC50 in Protheus.Sc5010s on new { Pedido = SD20.D2Pedido, Filial = SD20.D2Filial, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Pedido = SC50.C5Num, Filial = SC50.C5Filial, Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                                     where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" &&
                                     (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114)
                                     || ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf))
                                     && (int)(object)SD20.D2Emissao >= InicioAno && (int)(object)SD20.D2Emissao <= FimAno && SD20.D2Quant != 0
                                     select new
                                     {
                                         Filial = SD20.D2Filial,
                                         Cliente = SD20.D2Cliente,
                                         Loja = SD20.D2Loja,
                                         Nome = SA10.A1Nome,
                                         Tipo = SA10.A1Clinter,
                                         Est = SA10.A1Est,
                                         Mun = SA10.A1Mun,
                                         NF = SD20.D2Doc,
                                         Serie = SD20.D2Serie,
                                         Emissao = SD20.D2Emissao,
                                         Total = SD20.D2Total,
                                         Valipi = SD20.D2Valipi,
                                         Valicm = SD20.D2Valicm,
                                         Descon = SD20.D2Descon,
                                         TotalBrut = SD20.D2Valbrut
                                     })
                                            .GroupBy(x => new
                                            {
                                                x.Filial,
                                                x.Cliente,
                                                x.Loja,
                                                x.Nome,
                                                x.Tipo,
                                                x.Est,
                                                x.Mun,
                                                x.NF,
                                                x.Serie,
                                                x.Emissao
                                            })
                                            .Select(x => new RegistrosNF
                                            {
                                                Filial = x.Key.Filial,
                                                CliFor = x.Key.Cliente,
                                                Loja = x.Key.Loja,
                                                Nome = x.Key.Nome,
                                                Tipo = x.Key.Tipo == "H" ? "HOSPITAL" : x.Key.Tipo == "M" ? "MEDICO" : x.Key.Tipo == "I" ? "INSTRUMENTADOR" : x.Key.Tipo == "N" ? "NORMAL" : x.Key.Tipo == "C" ? "CONVENIO" : x.Key.Tipo == "P" ? "PARTICULAR" : x.Key.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                                NF = x.Key.NF,
                                                Serie = x.Key.Serie,
                                                Emissao = $"{x.Key.Emissao.Substring(6, 2)}/{x.Key.Emissao.Substring(4, 2)}/{x.Key.Emissao.Substring(0, 4)}",
                                                Total = x.Sum(c => c.TotalBrut),
                                                Valipi = x.Sum(c => c.Valipi),
                                                Valicm = x.Sum(c => c.Valicm),
                                                Descon = x.Sum(c => c.Descon),
                                                Mes = x.Key.Emissao.Substring(4, 2)
                                            }).ToList();

                var TipoFaturados = queryFaturado.Select(x => x.Tipo).Distinct().ToList();
                queryFaturado.ForEach(c =>
                {

                    if (!Faturamento.Any(x => x.Tipo == c.Tipo))
                    {
                        var Todos = queryFaturado.Where(x => x.Tipo == c.Tipo).ToList();

                        var NovoRelatorio = new RelatorioFaturamentoNF { Tipo = c.Tipo };

                        Todos.ForEach(d =>
                        {
                            switch (d.Mes)
                            {
                                case "01": NovoRelatorio.Janeiro += d.Total.Value; break;
                                case "02": NovoRelatorio.Fevereiro += d.Total.Value; break;
                                case "03": NovoRelatorio.Marco += d.Total.Value; break;
                                case "04": NovoRelatorio.Abril += d.Total.Value; break;
                                case "05": NovoRelatorio.Maio += d.Total.Value; break;
                                case "06": NovoRelatorio.Junho += d.Total.Value; break;
                                case "07": NovoRelatorio.Julho += d.Total.Value; break;
                                case "08": NovoRelatorio.Agosto += d.Total.Value; break;
                                case "09": NovoRelatorio.Setembro += d.Total.Value; break;
                                case "10": NovoRelatorio.Outubro += d.Total.Value; break;
                                case "11": NovoRelatorio.Novembro += d.Total.Value; break;
                                case "12": NovoRelatorio.Dezembro += d.Total.Value; break;
                            }

                            NovoRelatorio.Total += d.Total.Value;
                        });

                        Faturamento.Add(NovoRelatorio);
                    }
                });
                Faturamento = Faturamento.OrderBy(x => x.Tipo).ToList();
                #endregion

                #region Devolucao
                var queryDevolucao = (from SD10 in Protheus.Sd1010s
                                      join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, NF = SD10.D1Nfori, Serie = SD10.D1Seriori, Forne = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, NF = SF20.F2Doc, Serie = SF20.F2Serie, Forne = SF20.F2Cliente, Loja = SF20.F2Loja } into Sr
                                      from A in Sr.DefaultIfEmpty()
                                      join SA10 in Protheus.Sa1010s on new { Forne = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forne = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                      where SD10.DELET != "*" && A.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*"
                                      && CfNe.Contains((int)(object)SD10.D1Cf) && (int)(object)SD10.D1Dtdigit >= InicioAno && (int)(object)SD10.D1Dtdigit <= FimAno
                                      select new
                                      {
                                          Filial = SD10.D1Filial,
                                          Cliente = SD10.D1Fornece,
                                          Loja = SD10.D1Loja,
                                          Nome = SA10.A1Nome,
                                          Tipo = SA10.A1Clinter,
                                          Est = SA10.A1Est,
                                          Mun = SA10.A1Mun,
                                          NF = SD10.D1Doc,
                                          Serie = SD10.D1Serie,
                                          Emissao = SD10.D1Emissao,
                                          Total = SD10.D1Total - SD10.D1Valdesc,
                                          Valipi = SD10.D1Valipi,
                                          Valicm = SD10.D1Valicm,
                                          Descon = SD10.D1Valdesc,
                                          DTDIGIT = SD10.D1Dtdigit,
                                          TotalBruto = SD10.D1Total - SD10.D1Valdesc + SD10.D1Valipi
                                      }
                                         ).GroupBy(x => new
                                         {
                                             x.Filial,
                                             x.Cliente,
                                             x.Loja,
                                             x.Nome,
                                             x.Tipo,
                                             x.Est,
                                             x.Mun,
                                             x.NF,
                                             x.Serie,
                                             x.Emissao,
                                             x.DTDIGIT
                                         }).Select(x => new RegistrosNF
                                         {

                                             Filial = x.Key.Filial,
                                             CliFor = x.Key.Cliente,
                                             Loja = x.Key.Loja,
                                             Nome = x.Key.Nome,
                                             Tipo = x.Key.Tipo == "H" ? "HOSPITAL" : x.Key.Tipo == "M" ? "MEDICO" : x.Key.Tipo == "I" ? "INSTRUMENTADOR" : x.Key.Tipo == "N" ? "NORMAL" : x.Key.Tipo == "C" ? "CONVENIO" : x.Key.Tipo == "P" ? "PARTICULAR" : x.Key.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                             NF = x.Key.NF,
                                             Serie = x.Key.Serie,
                                             Emissao = $"{x.Key.Emissao.Substring(6, 2)}/{x.Key.Emissao.Substring(4, 2)}/{x.Key.Emissao.Substring(0, 4)}",
                                             Total = -x.Sum(c => c.TotalBruto),
                                             Valipi = -x.Sum(c => c.Valipi),
                                             Valicm = -x.Sum(c => c.Valicm),
                                             Descon = -x.Sum(c => c.Descon),
                                             Mes = x.Key.DTDIGIT.Substring(4, 2)
                                         }).ToList();


                queryDevolucao.ForEach(c =>
                {

                    if (!Devolucao.Any(x => x.Tipo == c.Tipo))
                    {
                        var Todos = queryDevolucao.Where(x => x.Tipo == c.Tipo).ToList();

                        var NovoRelatorio = new RelatorioDevolucaoNF { Tipo = c.Tipo };

                        Todos.ForEach(d =>
                        {
                            switch (d.Mes)
                            {
                                case "01": NovoRelatorio.Janeiro += d.Total.Value; break;
                                case "02": NovoRelatorio.Fevereiro += d.Total.Value; break;
                                case "03": NovoRelatorio.Marco += d.Total.Value; break;
                                case "04": NovoRelatorio.Abril += d.Total.Value; break;
                                case "05": NovoRelatorio.Maio += d.Total.Value; break;
                                case "06": NovoRelatorio.Junho += d.Total.Value; break;
                                case "07": NovoRelatorio.Julho += d.Total.Value; break;
                                case "08": NovoRelatorio.Agosto += d.Total.Value; break;
                                case "09": NovoRelatorio.Setembro += d.Total.Value; break;
                                case "10": NovoRelatorio.Outubro += d.Total.Value; break;
                                case "11": NovoRelatorio.Novembro += d.Total.Value; break;
                                case "12": NovoRelatorio.Dezembro += d.Total.Value; break;
                            }

                            NovoRelatorio.Total += d.Total.Value;
                        });

                        Devolucao.Add(NovoRelatorio);
                    }
                });

                Devolucao = Devolucao.OrderBy(x => x.Tipo).ToList();
                #endregion

                #region Relatorio
                var query = queryFaturado.Concat(queryDevolucao).ToList();

                var Tipos = query.Select(x => x.Tipo).Distinct().ToList();

                query.ForEach(c =>
                {

                    if (!Relatorio.Any(x => x.Tipo == c.Tipo))
                    {
                        var Todos = query.Where(x => x.Tipo == c.Tipo).ToList();

                        var NovoRelatorio = new RelatorioNF { Tipo = c.Tipo };

                        Todos.ForEach(d =>
                        {
                            switch (d.Mes)
                            {
                                case "01": NovoRelatorio.Janeiro += d.Total.Value; break;
                                case "02": NovoRelatorio.Fevereiro += d.Total.Value; break;
                                case "03": NovoRelatorio.Marco += d.Total.Value; break;
                                case "04": NovoRelatorio.Abril += d.Total.Value; break;
                                case "05": NovoRelatorio.Maio += d.Total.Value; break;
                                case "06": NovoRelatorio.Junho += d.Total.Value; break;
                                case "07": NovoRelatorio.Julho += d.Total.Value; break;
                                case "08": NovoRelatorio.Agosto += d.Total.Value; break;
                                case "09": NovoRelatorio.Setembro += d.Total.Value; break;
                                case "10": NovoRelatorio.Outubro += d.Total.Value; break;
                                case "11": NovoRelatorio.Novembro += d.Total.Value; break;
                                case "12": NovoRelatorio.Dezembro += d.Total.Value; break;
                            }

                            NovoRelatorio.Total += d.Total.Value;
                        });

                        Relatorio.Add(NovoRelatorio);
                    }

                });

                Relatorio = Relatorio.OrderBy(x => x.Tipo).ToList();
                #endregion


                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Faturamento NF Bruto Inter");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Faturamento NF Bruto Inter");

                sheet.Cells[1, 1].Value = "Tipo";
                sheet.Cells[1, 2].Value = "Janeiro";
                sheet.Cells[1, 3].Value = "Fevereiro";
                sheet.Cells[1, 4].Value = "Março";
                sheet.Cells[1, 5].Value = "Abril";
                sheet.Cells[1, 6].Value = "Maio";
                sheet.Cells[1, 7].Value = "Junho";
                sheet.Cells[1, 8].Value = "Julho";
                sheet.Cells[1, 9].Value = "Agosto";
                sheet.Cells[1, 10].Value = "Setembro";
                sheet.Cells[1, 11].Value = "Outubro";
                sheet.Cells[1, 12].Value = "Novembro";
                sheet.Cells[1, 13].Value = "Dezembro";
                sheet.Cells[1, 14].Value = "Total";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Tipo;
                    sheet.Cells[i, 2].Value = Pedido.Janeiro;
                    sheet.Cells[i, 3].Value = Pedido.Fevereiro;
                    sheet.Cells[i, 4].Value = Pedido.Marco;
                    sheet.Cells[i, 5].Value = Pedido.Abril;
                    sheet.Cells[i, 6].Value = Pedido.Maio;
                    sheet.Cells[i, 7].Value = Pedido.Junho;
                    sheet.Cells[i, 8].Value = Pedido.Julho;
                    sheet.Cells[i, 9].Value = Pedido.Agosto;
                    sheet.Cells[i, 10].Value = Pedido.Setembro;
                    sheet.Cells[i, 11].Value = Pedido.Outubro;
                    sheet.Cells[i, 12].Value = Pedido.Novembro;
                    sheet.Cells[i, 13].Value = Pedido.Dezembro;
                    sheet.Cells[i, 14].Value = Pedido.Total;

                    i++;
                });

                sheet.Cells[i, 1].Value = "TOTAL:";
                sheet.Cells[i, 2].Value = Relatorio.Sum(x => x.Janeiro);
                sheet.Cells[i, 3].Value = Relatorio.Sum(x => x.Fevereiro);
                sheet.Cells[i, 4].Value = Relatorio.Sum(x => x.Marco);
                sheet.Cells[i, 5].Value = Relatorio.Sum(x => x.Abril);
                sheet.Cells[i, 6].Value = Relatorio.Sum(x => x.Maio);
                sheet.Cells[i, 7].Value = Relatorio.Sum(x => x.Junho);
                sheet.Cells[i, 8].Value = Relatorio.Sum(x => x.Julho);
                sheet.Cells[i, 9].Value = Relatorio.Sum(x => x.Agosto);
                sheet.Cells[i, 10].Value = Relatorio.Sum(x => x.Setembro);
                sheet.Cells[i, 11].Value = Relatorio.Sum(x => x.Outubro);
                sheet.Cells[i, 12].Value = Relatorio.Sum(x => x.Novembro);
                sheet.Cells[i, 13].Value = Relatorio.Sum(x => x.Dezembro);
                sheet.Cells[i, 14].Value = Relatorio.Sum(x => x.Total);

                i++;
                i++;

                sheet.Cells[i, 1].Value = "NF Faturamento";
                sheet.Cells[i, 1, i, 14].Merge = true;
                i++;

                sheet.Cells[i, 1].Value = "Tipo";
                sheet.Cells[i, 2].Value = "Janeiro";
                sheet.Cells[i, 3].Value = "Fevereiro";
                sheet.Cells[i, 4].Value = "Março";
                sheet.Cells[i, 5].Value = "Abril";
                sheet.Cells[i, 6].Value = "Maio";
                sheet.Cells[i, 7].Value = "Junho";
                sheet.Cells[i, 8].Value = "Julho";
                sheet.Cells[i, 9].Value = "Agosto";
                sheet.Cells[i, 10].Value = "Setembro";
                sheet.Cells[i, 11].Value = "Outubro";
                sheet.Cells[i, 12].Value = "Novembro";
                sheet.Cells[i, 13].Value = "Dezembro";
                sheet.Cells[i, 14].Value = "Total";
                i++;

                Faturamento.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Tipo;
                    sheet.Cells[i, 2].Value = Pedido.Janeiro;
                    sheet.Cells[i, 3].Value = Pedido.Fevereiro;
                    sheet.Cells[i, 4].Value = Pedido.Marco;
                    sheet.Cells[i, 5].Value = Pedido.Abril;
                    sheet.Cells[i, 6].Value = Pedido.Maio;
                    sheet.Cells[i, 7].Value = Pedido.Junho;
                    sheet.Cells[i, 8].Value = Pedido.Julho;
                    sheet.Cells[i, 9].Value = Pedido.Agosto;
                    sheet.Cells[i, 10].Value = Pedido.Setembro;
                    sheet.Cells[i, 11].Value = Pedido.Outubro;
                    sheet.Cells[i, 12].Value = Pedido.Novembro;
                    sheet.Cells[i, 13].Value = Pedido.Dezembro;
                    sheet.Cells[i, 14].Value = Pedido.Total;

                    i++;
                });

                sheet.Cells[i, 1].Value = "TOTAL:";
                sheet.Cells[i, 2].Value = Faturamento.Sum(x => x.Janeiro);
                sheet.Cells[i, 3].Value = Faturamento.Sum(x => x.Fevereiro);
                sheet.Cells[i, 4].Value = Faturamento.Sum(x => x.Marco);
                sheet.Cells[i, 5].Value = Faturamento.Sum(x => x.Abril);
                sheet.Cells[i, 6].Value = Faturamento.Sum(x => x.Maio);
                sheet.Cells[i, 7].Value = Faturamento.Sum(x => x.Junho);
                sheet.Cells[i, 8].Value = Faturamento.Sum(x => x.Julho);
                sheet.Cells[i, 9].Value = Faturamento.Sum(x => x.Agosto);
                sheet.Cells[i, 10].Value = Faturamento.Sum(x => x.Setembro);
                sheet.Cells[i, 11].Value = Faturamento.Sum(x => x.Outubro);
                sheet.Cells[i, 12].Value = Faturamento.Sum(x => x.Novembro);
                sheet.Cells[i, 13].Value = Faturamento.Sum(x => x.Dezembro);
                sheet.Cells[i, 14].Value = Faturamento.Sum(x => x.Total);

                i++;
                i++;


                sheet.Cells[i, 1].Value = "NF Devolução";
                sheet.Cells[i, 1, i, 14].Merge = true;
                i++;

                sheet.Cells[i, 1].Value = "Tipo";
                sheet.Cells[i, 2].Value = "Janeiro";
                sheet.Cells[i, 3].Value = "Fevereiro";
                sheet.Cells[i, 4].Value = "Março";
                sheet.Cells[i, 5].Value = "Abril";
                sheet.Cells[i, 6].Value = "Maio";
                sheet.Cells[i, 7].Value = "Junho";
                sheet.Cells[i, 8].Value = "Julho";
                sheet.Cells[i, 9].Value = "Agosto";
                sheet.Cells[i, 10].Value = "Setembro";
                sheet.Cells[i, 11].Value = "Outubro";
                sheet.Cells[i, 12].Value = "Novembro";
                sheet.Cells[i, 13].Value = "Dezembro";
                sheet.Cells[i, 14].Value = "Total";
                i++;

                Devolucao.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Tipo;
                    sheet.Cells[i, 2].Value = Pedido.Janeiro;
                    sheet.Cells[i, 3].Value = Pedido.Fevereiro;
                    sheet.Cells[i, 4].Value = Pedido.Marco;
                    sheet.Cells[i, 5].Value = Pedido.Abril;
                    sheet.Cells[i, 6].Value = Pedido.Maio;
                    sheet.Cells[i, 7].Value = Pedido.Junho;
                    sheet.Cells[i, 8].Value = Pedido.Julho;
                    sheet.Cells[i, 9].Value = Pedido.Agosto;
                    sheet.Cells[i, 10].Value = Pedido.Setembro;
                    sheet.Cells[i, 11].Value = Pedido.Outubro;
                    sheet.Cells[i, 12].Value = Pedido.Novembro;
                    sheet.Cells[i, 13].Value = Pedido.Dezembro;
                    sheet.Cells[i, 14].Value = Pedido.Total;

                    i++;
                });

                sheet.Cells[i, 1].Value = "TOTAL:";
                sheet.Cells[i, 2].Value = Devolucao.Sum(x => x.Janeiro);
                sheet.Cells[i, 3].Value = Devolucao.Sum(x => x.Fevereiro);
                sheet.Cells[i, 4].Value = Devolucao.Sum(x => x.Marco);
                sheet.Cells[i, 5].Value = Devolucao.Sum(x => x.Abril);
                sheet.Cells[i, 6].Value = Devolucao.Sum(x => x.Maio);
                sheet.Cells[i, 7].Value = Devolucao.Sum(x => x.Junho);
                sheet.Cells[i, 8].Value = Devolucao.Sum(x => x.Julho);
                sheet.Cells[i, 9].Value = Devolucao.Sum(x => x.Agosto);
                sheet.Cells[i, 10].Value = Devolucao.Sum(x => x.Setembro);
                sheet.Cells[i, 11].Value = Devolucao.Sum(x => x.Outubro);
                sheet.Cells[i, 12].Value = Devolucao.Sum(x => x.Novembro);
                sheet.Cells[i, 13].Value = Devolucao.Sum(x => x.Dezembro);
                sheet.Cells[i, 14].Value = Devolucao.Sum(x => x.Total);


                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FaturamentoNFBRUTOInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoNFBRUTOInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
