using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Controladoria.FaturamentoNF;
using SGID.Models.Controladoria.FaturamentoNF.RegistrosFaturamento;
using SGID.Models.Relatorio;

namespace SGID.Pages.Relatorios.Controladoria
{
    [Authorize]
    public class FaturamentoNFBRUTOModel : PageModel
    {
        private RelatorioContext DB { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioNF> Relatorio { get; set; } = new List<RelatorioNF>();
        public List<RelatorioFaturamentoNF> Faturamento { get; set; } = new List<RelatorioFaturamentoNF>();
        public List<RelatorioDevolucaoNF> Devolucao { get; set; } = new List<RelatorioDevolucaoNF>();

        public string Ano { get; set; }

        public FaturamentoNFBRUTOModel(RelatorioContext dB,ApplicationDbContext sgid)
        {
            DB = dB;
            SGID = sgid;
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

                #region Relatorio
                var query = DB.Fatnfs.Where(x => x.Empresa == "D" && ((int)(object)x.Emissao >= InicioAno && (int)(object)x.Emissao <= FimAno || (int)(object)x.Dtdigit >= InicioAno && (int)(object)x.Dtdigit <= FimAno))
                    .Select(x => new
                    {
                        Empresa = x.Empresa,
                        Filial = x.Filial,
                        Tipo = x.Tipo == "H" ? "HOSPITAL" : x.Tipo == "G" ? "INTERGRUPO" : x.Tipo == "M" ? "MEDICO" : x.Tipo == "I" ? "INSTRUMENTADOR" : x.Tipo == "N" ? "NORMAL" : x.Tipo == "C" ? "CONVENIO" : x.Tipo == "P" ? "PARTICULAR" : x.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                        Mes = x.Tipofd == "F" ? x.Emissao.Substring(4, 2) : x.Dtdigit.Substring(4, 2),
                        Total = x.Totalbrut
                    })
                    .GroupBy(x => new
                    {
                        x.Empresa,
                        x.Filial,
                        x.Tipo,
                        x.Mes
                    })
                    .Select(x => new
                    {
                        x.Key.Empresa,
                        x.Key.Filial,
                        x.Key.Tipo,
                        x.Key.Mes,
                        Total = x.Sum(c => c.Total)
                    })
                    .ToList();

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

                #region Faturado

                var queryFaturado = DB.Fatnfs.Where(x => x.Empresa == "D" && x.Tipofd == "F" && (int)(object)x.Emissao >= InicioAno && (int)(object)x.Emissao <= FimAno)
                                             .Select(x => new RegistrosNF
                                             {
                                                 Empresa = x.Empresa,
                                                 Filial = x.Filial,
                                                 CliFor = x.Clifor,
                                                 Loja = x.Loja,
                                                 Nome = x.Nome,
                                                 Tipo = x.Tipo == "H" ? "HOSPITAL" : x.Tipo == "G" ? "INTERGRUPO" : x.Tipo == "M" ? "MEDICO" : x.Tipo == "I" ? "INSTRUMENTADOR" : x.Tipo == "N" ? "NORMAL" : x.Tipo == "C" ? "CONVENIO" : x.Tipo == "P" ? "PARTICULAR" : x.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                                 NF = x.Nf,
                                                 Serie = x.Serie,
                                                 Emissao = $"{x.Emissao.Substring(6, 2)}/{x.Emissao.Substring(4, 2)}/{x.Emissao.Substring(0, 4)}",
                                                 Total = x.Totalbrut,
                                                 Valipi = x.Valipi,
                                                 Valicm = x.Valicm,
                                                 Descon = x.Descon,
                                                 Mes = x.Emissao.Substring(4, 2)
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

                var queryDevolucao = DB.Fatnfs.Where(x => x.Empresa == "D" && x.Tipofd == "D" && (int)(object)x.Emissao >= InicioAno && (int)(object)x.Emissao <= FimAno)
                                             .Select(x => new RegistrosNF
                                             {
                                                 Empresa = x.Empresa,
                                                 Filial = x.Filial,
                                                 CliFor = x.Clifor,
                                                 Loja = x.Loja,
                                                 Nome = x.Nome,
                                                 Tipo = x.Tipo == "H" ? "HOSPITAL"  : x.Tipo == "G" ? "INTERGRUPO" : x.Tipo == "M" ? "MEDICO" : x.Tipo == "I" ? "INSTRUMENTADOR" : x.Tipo == "N" ? "NORMAL" : x.Tipo == "C" ? "CONVENIO" : x.Tipo == "P" ? "PARTICULAR" : x.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                                 NF = x.Nf,
                                                 Serie = x.Serie,
                                                 Emissao = $"{x.Emissao.Substring(6, 2)}/{x.Emissao.Substring(4, 2)}/{x.Emissao.Substring(0, 4)}",
                                                 Total = x.Totalbrut,
                                                 Valipi = x.Valipi,
                                                 Valicm = x.Valicm,
                                                 Descon = x.Descon,
                                                 Mes = x.Emissao.Substring(4, 2)
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

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoNFBRUTO",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(string Ano)
        {
            try
            {

                var InicioAno = Convert.ToInt32($"{Ano}0101");
                var FimAno = Convert.ToInt32($"{Ano}1231");


                #region Relatorio
                var query = DB.Fatnfs.Where(x => x.Empresa == "D" && ((int)(object)x.Emissao >= InicioAno && (int)(object)x.Emissao <= FimAno || (int)(object)x.Dtdigit >= InicioAno && (int)(object)x.Dtdigit <= FimAno))
                    .Select(x => new
                    {
                        Empresa = x.Empresa,
                        Filial = x.Filial,
                        Tipo = x.Tipo == "H" ? "HOSPITAL" : x.Tipo == "G" ? "INTERGRUPO" : x.Tipo == "M" ? "MEDICO" : x.Tipo == "I" ? "INSTRUMENTADOR" : x.Tipo == "N" ? "NORMAL" : x.Tipo == "C" ? "CONVENIO" : x.Tipo == "P" ? "PARTICULAR" : x.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                        Mes = x.Tipofd == "F" ? x.Emissao.Substring(4, 2) : x.Dtdigit.Substring(4, 2),
                        Total = x.Totalbrut
                    })
                    .GroupBy(x => new
                    {
                        x.Empresa,
                        x.Filial,
                        x.Tipo,
                        x.Mes
                    })
                    .Select(x => new
                    {
                        x.Key.Empresa,
                        x.Key.Filial,
                        x.Key.Tipo,
                        x.Key.Mes,
                        Total = x.Sum(c => c.Total)
                    })
                    .ToList();

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

                #region Faturado

                var queryFaturado = DB.Fatnfs.Where(x => x.Empresa == "D" && x.Tipofd == "F" && (int)(object)x.Emissao >= InicioAno && (int)(object)x.Emissao <= FimAno)
                                             .Select(x => new RegistrosNF
                                             {
                                                 Empresa = x.Empresa,
                                                 Filial = x.Filial,
                                                 CliFor = x.Clifor,
                                                 Loja = x.Loja,
                                                 Nome = x.Nome,
                                                 Tipo = x.Tipo == "H" ? "HOSPITAL" : x.Tipo == "G" ? "INTERGRUPO" : x.Tipo == "M" ? "MEDICO" : x.Tipo == "I" ? "INSTRUMENTADOR" : x.Tipo == "N" ? "NORMAL" : x.Tipo == "C" ? "CONVENIO" : x.Tipo == "P" ? "PARTICULAR" : x.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                                 NF = x.Nf,
                                                 Serie = x.Serie,
                                                 Emissao = $"{x.Emissao.Substring(6, 2)}/{x.Emissao.Substring(4, 2)}/{x.Emissao.Substring(0, 4)}",
                                                 Total = x.Totalbrut,
                                                 Valipi = x.Valipi,
                                                 Valicm = x.Valicm,
                                                 Descon = x.Descon,
                                                 Mes = x.Emissao.Substring(4, 2)
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

                var queryDevolucao = DB.Fatnfs.Where(x => x.Empresa == "D" && x.Tipofd == "D" && (int)(object)x.Emissao >= InicioAno && (int)(object)x.Emissao <= FimAno)
                                             .Select(x => new RegistrosNF
                                             {
                                                 Empresa = x.Empresa,
                                                 Filial = x.Filial,
                                                 CliFor = x.Clifor,
                                                 Loja = x.Loja,
                                                 Nome = x.Nome,
                                                 Tipo = x.Tipo == "H" ? "HOSPITAL" : x.Tipo == "G" ? "INTERGRUPO" : x.Tipo == "M" ? "MEDICO" : x.Tipo == "I" ? "INSTRUMENTADOR" : x.Tipo == "N" ? "NORMAL" : x.Tipo == "C" ? "CONVENIO" : x.Tipo == "P" ? "PARTICULAR" : x.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                                 NF = x.Nf,
                                                 Serie = x.Serie,
                                                 Emissao = $"{x.Emissao.Substring(6, 2)}/{x.Emissao.Substring(4, 2)}/{x.Emissao.Substring(0, 4)}",
                                                 Total = x.Totalbrut,
                                                 Valipi = x.Valipi,
                                                 Valicm = x.Valicm,
                                                 Descon = x.Descon,
                                                 Mes = x.Emissao.Substring(4, 2)
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


                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Faturamento NF Bruto");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Faturamento NF Bruto");

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
                sheet.Cells[i, 2].Value = Relatorio.Sum(x=>x.Janeiro);
                sheet.Cells[i, 3].Value = Relatorio.Sum(x=>x.Fevereiro);
                sheet.Cells[i, 4].Value = Relatorio.Sum(x=>x.Marco);
                sheet.Cells[i, 5].Value = Relatorio.Sum(x=>x.Abril);
                sheet.Cells[i, 6].Value = Relatorio.Sum(x=>x.Maio);
                sheet.Cells[i, 7].Value = Relatorio.Sum(x=>x.Junho);
                sheet.Cells[i, 8].Value = Relatorio.Sum(x=>x.Julho);
                sheet.Cells[i, 9].Value = Relatorio.Sum(x => x.Agosto);
                sheet.Cells[i, 10].Value = Relatorio.Sum(x=>x.Setembro);
                sheet.Cells[i, 11].Value = Relatorio.Sum(x=>x.Outubro);
                sheet.Cells[i, 12].Value = Relatorio.Sum(x=>x.Novembro);
                sheet.Cells[i, 13].Value = Relatorio.Sum(x=>x.Dezembro);
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

                i++;

                //var format = sheet.Cells[2, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FaturamentoNFBRUTO.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoNFBRUTO Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
