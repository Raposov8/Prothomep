using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Estoque.RelatorioFaturamentoNFFab;
using SGID.Models.Relatorio;
using System.Linq;

namespace SGID.Pages.Relatorios.Estoque
{
    [Authorize]
    public class FaturamentoNFFabModel : PageModel
    {
        private RelatorioContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<Tipo> RelatorioFaturamento { get; set; } = new List<Tipo>();

        public List<Tipo> RelatorioDevolucao { get; set; } = new List<Tipo>();

        public List<Fabricante> RelatorioFabricante { get; set; } = new List<Fabricante>();

        public FaturamentoNFFabModel(RelatorioContext context,ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        public string Ano { get; set; }
        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string Ano)
        {
            try
            {
                this.Ano = Ano;

                int Inicio = Convert.ToInt32($"{Ano}0101");
                int Fim = Convert.ToInt32($"{Ano}1231");

                #region Fabricante
                var query = Protheus.Fatnffabs.Where(x => x.Empresa == "D" && (((int)(object)x.Emissao >= Inicio && (int)(object)x.Emissao <= Fim && x.Tipofd == "F")
                || ((int)(object)x.Dtdigit >= Inicio && (int)(object)x.Dtdigit <= Fim && x.Tipofd != "F")))
                    .
                    Select(novo => new
                    {
                        novo.Empresa,
                        novo.Filial,
                        novo.Fabricante,
                        Mest = novo.Tipofd == "F" ? novo.Emissao.Substring(4, 2) : novo.Dtdigit.Substring(4, 2),
                        novo.Total
                    })
                    .GroupBy(x => new
                    {
                        x.Empresa,
                        x.Filial,
                        x.Fabricante,
                        x.Mest
                    })
                    .Select(x => new
                    {
                        x.Key.Empresa,
                        x.Key.Filial,
                        x.Key.Fabricante,
                        x.Key.Mest,
                        Total = x.Sum(c => c.Total)
                    }).OrderBy(x => x.Fabricante).ToList();


                query.ForEach(c =>
                {

                    if (!RelatorioFabricante.Any(x => x.Nome == c.Fabricante))
                    {
                        var Todos = query.Where(x => x.Fabricante == c.Fabricante).ToList();

                        var NovoRelatorio = new Fabricante { Nome = c.Fabricante };

                        Todos.ForEach(x =>
                        {
                            switch (x.Mest)
                            {
                                case "01": NovoRelatorio.Janeiro += x.Total.Value; break;
                                case "02": NovoRelatorio.Fevereiro += x.Total.Value; break;
                                case "03": NovoRelatorio.Marco += x.Total.Value; break;
                                case "04": NovoRelatorio.Abril += x.Total.Value; break;
                                case "05": NovoRelatorio.Maio += x.Total.Value; break;
                                case "06": NovoRelatorio.Junho += x.Total.Value; break;
                                case "07": NovoRelatorio.Julho += x.Total.Value; break;
                                case "08": NovoRelatorio.Agosto += x.Total.Value; break;
                                case "09": NovoRelatorio.Setembro += x.Total.Value; break;
                                case "10": NovoRelatorio.Outubro += x.Total.Value; break;
                                case "11": NovoRelatorio.Novembro += x.Total.Value; break;
                                case "12": NovoRelatorio.Dezembro += x.Total.Value; break;
                            }

                            NovoRelatorio.Total += x.Total.Value;
                        });

                        RelatorioFabricante.Add(NovoRelatorio);
                    }

                });

                #endregion

                #region Faturado

                var queryFaturado = Protheus.Fatnffabs.Where(x => x.Empresa == "D" && x.Tipofd == "F"
                && (int)(object)x.Emissao >= Inicio && (int)(object)x.Emissao <= Fim)
                    .Select(x => new
                    {
                        x.Empresa,
                        x.Filial,
                        x.Clifor,
                        x.Loja,
                        x.Nome,
                        Tipo = x.Tipo == "H" ? "HOSPITAL" : x.Tipo == "M" ? "MEDICO" : x.Tipo == "I" ? "INSTRUMENTAL" : x.Tipo == "N" ? "NORMAL" : x.Tipo == "C" ? "CONVENIO" : x.Tipo == "P" ? "PARTICULAR" : x.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                        x.Est,
                        x.Mun,
                        x.Nf,
                        x.Serie,
                        x.Emissao,
                        x.Fabricante,
                        x.Total,
                        x.Valipi,
                        x.Valicm,
                        x.Descon,
                        Mes = x.Emissao.Substring(4, 2)
                    }).OrderBy(x => x.Fabricante).ToList();


                var hospital = queryFaturado.Where(x => x.Tipo == "HOSPITAL").ToList();
                var Medico = queryFaturado.Where(x => x.Tipo == "MEDICO").ToList();
                var Instrumental = queryFaturado.Where(x => x.Tipo == "INSTRUMENTAL").ToList();
                var Normal = queryFaturado.Where(x => x.Tipo == "NORMAL").ToList();
                var Convenio = queryFaturado.Where(x => x.Tipo == "CONVENIO").ToList();
                var Particular = queryFaturado.Where(x => x.Tipo == "PARTICULAR").ToList();
                var Sub = queryFaturado.Where(x => x.Tipo == "SUB-DISTRIBUIDOR").ToList();
                var Outros = queryFaturado.Where(x => x.Tipo == "OUTROS").ToList();

                if (hospital.Count > 0)
                {
                    var tipo = new Tipo { Nome = "HOSPITAL" };

                    hospital.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = hospital.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Medico.Count > 0)
                {
                    var tipo = new Tipo { Nome = "MEDICO" };

                    Medico.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Medico.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Instrumental.Count > 0)
                {
                    var tipo = new Tipo { Nome = "INSTRUMENTAL" };

                    Instrumental.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Instrumental.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Normal.Count > 0)
                {
                    var tipo = new Tipo { Nome = "NORMAL" };

                    Normal.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Normal.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Convenio.Count > 0)
                {
                    var tipo = new Tipo { Nome = "CONVENIO" };

                    Convenio.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Convenio.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Particular.Count > 0)
                {
                    var tipo = new Tipo { Nome = "PARTICULAR" };

                    Particular.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Particular.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Sub.Count > 0)
                {
                    var tipo = new Tipo { Nome = "SUB-DISTRIBUIDOR" };

                    Sub.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Sub.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Outros.Count > 0)
                {
                    var tipo = new Tipo { Nome = "OUTROS" };

                    Outros.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Outros.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                #endregion

                #region Devolucao

                var queryDevolucao = Protheus.Fatnffabs.Where(x => x.Empresa == "D" && x.Tipofd == "D"
                && (int)(object)x.Emissao >= Inicio && (int)(object)x.Emissao <= Fim)
                    .Select(x => new
                    {
                        x.Empresa,
                        x.Filial,
                        x.Clifor,
                        x.Loja,
                        x.Nome,
                        Tipo = x.Tipo == "H" ? "HOSPITAL" : x.Tipo == "M" ? "MEDICO" : x.Tipo == "I" ? "INSTRUMENTAL" : x.Tipo == "N" ? "NORMAL" : x.Tipo == "C" ? "CONVENIO" : x.Tipo == "P" ? "PARTICULAR" : x.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                        x.Est,
                        x.Mun,
                        x.Nf,
                        x.Serie,
                        x.Emissao,
                        x.Fabricante,
                        x.Total,
                        x.Valipi,
                        x.Valicm,
                        x.Descon,
                        Mes = x.Emissao.Substring(4, 2)
                    }).OrderBy(x => x.Fabricante).ToList();

                var hospitalDev = queryDevolucao.Where(x => x.Tipo == "HOSPITAL").ToList();
                var MedicoDev = queryDevolucao.Where(x => x.Tipo == "MEDICO").ToList();
                var InstrumentalDev = queryDevolucao.Where(x => x.Tipo == "INSTRUMENTAL").ToList();
                var NormalDev = queryDevolucao.Where(x => x.Tipo == "NORMAL").ToList();
                var ConvenioDev = queryDevolucao.Where(x => x.Tipo == "CONVENIO").ToList();
                var ParticularDev = queryDevolucao.Where(x => x.Tipo == "PARTICULAR").ToList();
                var SubDev = queryDevolucao.Where(x => x.Tipo == "SUB-DISTRIBUIDOR").ToList();
                var OutrosDev = queryDevolucao.Where(x => x.Tipo == "OUTROS").ToList();

                if (hospitalDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "HOSPITAL" };

                    hospitalDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = hospitalDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (MedicoDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "MEDICO" };

                    MedicoDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = MedicoDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (InstrumentalDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "INSTRUMENTAL" };

                    InstrumentalDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = InstrumentalDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (NormalDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "NORMAL" };

                    NormalDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = NormalDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (ConvenioDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "CONVENIO" };

                    ConvenioDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = ConvenioDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (ParticularDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "PARTICULAR" };

                    ParticularDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = ParticularDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (SubDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "SUB-DISTRIBUIDOR" };

                    SubDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = SubDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (OutrosDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "OUTROS" };

                    OutrosDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = OutrosDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }

                #endregion

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoNFFAB",user);

                return LocalRedirect("/error");
            }
        }

        public IActionResult OnPostExport(string Ano)
        {
            try
            {
                int Inicio = Convert.ToInt32($"{Ano}0101");
                int Fim = Convert.ToInt32($"{Ano}1231");

                #region Fabricante
                var query = Protheus.Fatnffabs.Where(x => x.Empresa == "D" && (((int)(object)x.Emissao >= Inicio && (int)(object)x.Emissao <= Fim && x.Tipofd == "F")
                || ((int)(object)x.Dtdigit >= Inicio && (int)(object)x.Dtdigit <= Fim && x.Tipofd != "F")))
                    .
                    Select(novo => new
                    {
                        novo.Empresa,
                        novo.Filial,
                        novo.Fabricante,
                        Mest = novo.Tipofd == "F" ? novo.Emissao.Substring(4, 2) : novo.Dtdigit.Substring(4, 2),
                        novo.Total
                    })
                    .GroupBy(x => new
                    {
                        x.Empresa,
                        x.Filial,
                        x.Fabricante,
                        x.Mest
                    })
                    .Select(x => new
                    {
                        x.Key.Empresa,
                        x.Key.Filial,
                        x.Key.Fabricante,
                        x.Key.Mest,
                        Total = x.Sum(c => c.Total)
                    }).OrderBy(x => x.Fabricante).ToList();


                query.ForEach(c =>
                {

                    if (!RelatorioFabricante.Any(x => x.Nome == c.Fabricante))
                    {
                        var Todos = query.Where(x => x.Fabricante == c.Fabricante).ToList();

                        var NovoRelatorio = new Fabricante { Nome = c.Fabricante };

                        Todos.ForEach(x =>
                        {
                            switch (x.Mest)
                            {
                                case "01": NovoRelatorio.Janeiro += x.Total.Value; break;
                                case "02": NovoRelatorio.Fevereiro += x.Total.Value; break;
                                case "03": NovoRelatorio.Marco += x.Total.Value; break;
                                case "04": NovoRelatorio.Abril += x.Total.Value; break;
                                case "05": NovoRelatorio.Maio += x.Total.Value; break;
                                case "06": NovoRelatorio.Junho += x.Total.Value; break;
                                case "07": NovoRelatorio.Julho += x.Total.Value; break;
                                case "08": NovoRelatorio.Agosto += x.Total.Value; break;
                                case "09": NovoRelatorio.Setembro += x.Total.Value; break;
                                case "10": NovoRelatorio.Outubro += x.Total.Value; break;
                                case "11": NovoRelatorio.Novembro += x.Total.Value; break;
                                case "12": NovoRelatorio.Dezembro += x.Total.Value; break;
                            }

                            NovoRelatorio.Total += x.Total.Value;
                        });

                        RelatorioFabricante.Add(NovoRelatorio);
                    }

                });

                #endregion

                #region Faturado

                var queryFaturado = Protheus.Fatnffabs.Where(x => x.Empresa == "D" && x.Tipofd == "F"
                && (int)(object)x.Emissao >= Inicio && (int)(object)x.Emissao <= Fim)
                    .Select(x => new
                    {
                        x.Empresa,
                        x.Filial,
                        x.Clifor,
                        x.Loja,
                        x.Nome,
                        Tipo = x.Tipo == "H" ? "HOSPITAL" : x.Tipo == "M" ? "MEDICO" : x.Tipo == "I" ? "INSTRUMENTAL" : x.Tipo == "N" ? "NORMAL" : x.Tipo == "C" ? "CONVENIO" : x.Tipo == "P" ? "PARTICULAR" : x.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                        x.Est,
                        x.Mun,
                        x.Nf,
                        x.Serie,
                        x.Emissao,
                        x.Fabricante,
                        x.Total,
                        x.Valipi,
                        x.Valicm,
                        x.Descon,
                        Mes = x.Emissao.Substring(4, 2)
                    }).OrderBy(x => x.Fabricante).ToList();


                var hospital = queryFaturado.Where(x => x.Tipo == "HOSPITAL").ToList();
                var Medico = queryFaturado.Where(x => x.Tipo == "MEDICO").ToList();
                var Instrumental = queryFaturado.Where(x => x.Tipo == "INSTRUMENTAL").ToList();
                var Normal = queryFaturado.Where(x => x.Tipo == "NORMAL").ToList();
                var Convenio = queryFaturado.Where(x => x.Tipo == "CONVENIO").ToList();
                var Particular = queryFaturado.Where(x => x.Tipo == "PARTICULAR").ToList();
                var Sub = queryFaturado.Where(x => x.Tipo == "SUB-DISTRIBUIDOR").ToList();
                var Outros = queryFaturado.Where(x => x.Tipo == "OUTROS").ToList();

                if (hospital.Count > 0)
                {
                    var tipo = new Tipo { Nome = "HOSPITAL" };

                    hospital.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = hospital.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Medico.Count > 0)
                {
                    var tipo = new Tipo { Nome = "MEDICO" };

                    Medico.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Medico.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Instrumental.Count > 0)
                {
                    var tipo = new Tipo { Nome = "INSTRUMENTAL" };

                    Instrumental.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Instrumental.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Normal.Count > 0)
                {
                    var tipo = new Tipo { Nome = "NORMAL" };

                    Normal.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Normal.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Convenio.Count > 0)
                {
                    var tipo = new Tipo { Nome = "CONVENIO" };

                    Convenio.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Convenio.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Particular.Count > 0)
                {
                    var tipo = new Tipo { Nome = "PARTICULAR" };

                    Particular.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Particular.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Sub.Count > 0)
                {
                    var tipo = new Tipo { Nome = "SUB-DISTRIBUIDOR" };

                    Sub.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Sub.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                if (Outros.Count > 0)
                {
                    var tipo = new Tipo { Nome = "OUTROS" };

                    Outros.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = Outros.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }

                RelatorioFaturamento = RelatorioFaturamento.OrderBy(x => x.Nome).ToList();
                #endregion

                #region Devolucao

                var queryDevolucao = Protheus.Fatnffabs.Where(x => x.Empresa == "D" && x.Tipofd == "D"
                && (int)(object)x.Emissao >= Inicio && (int)(object)x.Emissao <= Fim)
                    .Select(x => new
                    {
                        x.Empresa,
                        x.Filial,
                        x.Clifor,
                        x.Loja,
                        x.Nome,
                        Tipo = x.Tipo == "H" ? "HOSPITAL" : x.Tipo == "M" ? "MEDICO" : x.Tipo == "I" ? "INSTRUMENTAL" : x.Tipo == "N" ? "NORMAL" : x.Tipo == "C" ? "CONVENIO" : x.Tipo == "P" ? "PARTICULAR" : x.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                        x.Est,
                        x.Mun,
                        x.Nf,
                        x.Serie,
                        x.Emissao,
                        x.Fabricante,
                        x.Total,
                        x.Valipi,
                        x.Valicm,
                        x.Descon,
                        Mes = x.Emissao.Substring(4, 2)
                    }).OrderBy(x => x.Fabricante).ToList();

                var hospitalDev = queryDevolucao.Where(x => x.Tipo == "HOSPITAL").ToList();
                var MedicoDev = queryDevolucao.Where(x => x.Tipo == "MEDICO").ToList();
                var InstrumentalDev = queryDevolucao.Where(x => x.Tipo == "INSTRUMENTAL").ToList();
                var NormalDev = queryDevolucao.Where(x => x.Tipo == "NORMAL").ToList();
                var ConvenioDev = queryDevolucao.Where(x => x.Tipo == "CONVENIO").ToList();
                var ParticularDev = queryDevolucao.Where(x => x.Tipo == "PARTICULAR").ToList();
                var SubDev = queryDevolucao.Where(x => x.Tipo == "SUB-DISTRIBUIDOR").ToList();
                var OutrosDev = queryDevolucao.Where(x => x.Tipo == "OUTROS").ToList();

                if (hospitalDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "HOSPITAL" };

                    hospitalDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = hospitalDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (MedicoDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "MEDICO" };

                    MedicoDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = MedicoDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (InstrumentalDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "INSTRUMENTAL" };

                    InstrumentalDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = InstrumentalDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (NormalDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "NORMAL" };

                    NormalDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = NormalDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (ConvenioDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "CONVENIO" };

                    ConvenioDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = ConvenioDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (ParticularDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "PARTICULAR" };

                    ParticularDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = ParticularDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (SubDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "SUB-DISTRIBUIDOR" };

                    SubDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = SubDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }
                if (OutrosDev.Count > 0)
                {
                    var tipo = new Tipo { Nome = "OUTROS" };

                    OutrosDev.ForEach(x =>
                    {
                        if (!tipo.Clientes.Any(c => c.Nome == x.Fabricante))
                        {
                            var cliente = new Cliente { Nome = x.Fabricante };

                            var notas = OutrosDev.Where(c => c.Fabricante == cliente.Nome).ToList();

                            notas.ForEach(nota =>
                            {


                                var todos = notas.Where(x => x.Nome == nota.Nome).ToList();

                                var NotaFiscal = new NotaFiscal
                                {
                                    Nome = nota.Nome,
                                    Serie = nota.Serie,
                                    NF = nota.Nf,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total.Value;
                                            NotaFiscal.Janeiro += nota.Total.Value; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total.Value;
                                            NotaFiscal.Fevereiro += nota.Total.Value; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total.Value;
                                            NotaFiscal.Marco += nota.Total.Value; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total.Value;
                                            NotaFiscal.Abril += nota.Total.Value; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total.Value;
                                            NotaFiscal.Maio += nota.Total.Value; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total.Value;
                                            NotaFiscal.Junho += nota.Total.Value; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total.Value;
                                            NotaFiscal.Julho += nota.Total.Value; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total.Value;
                                            NotaFiscal.Agosto += nota.Total.Value; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total.Value;
                                            NotaFiscal.Setembro += nota.Total.Value; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total.Value;
                                            NotaFiscal.Outubro += nota.Total.Value; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total.Value;
                                            NotaFiscal.Novembro += nota.Total.Value; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total.Value;
                                            NotaFiscal.Dezembro += nota.Total.Value; break;
                                        }
                                }

                                tipo.Total += x.Total.Value;
                                NotaFiscal.Total += x.Total.Value;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }

                RelatorioDevolucao = RelatorioDevolucao.OrderBy(x => x.Nome).ToList();

                #endregion

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Faturamento NF Fab");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Faturamento NF Fab");

                sheet.Cells[1, 1].Value = "Fabricante";
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


                int i = 2;

                RelatorioFabricante.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Nome;
                    sheet.Cells[i, 2].Value = Pedido.Janeiro;
                    sheet.Cells[i, 2].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 3].Value = Pedido.Fevereiro;
                    sheet.Cells[i, 3].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 4].Value = Pedido.Marco;
                    sheet.Cells[i, 4].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 5].Value = Pedido.Abril;
                    sheet.Cells[i, 5].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 6].Value = Pedido.Maio;
                    sheet.Cells[i, 6].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 7].Value = Pedido.Junho;
                    sheet.Cells[i, 7].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 8].Value = Pedido.Julho;
                    sheet.Cells[i, 8].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 9].Value = Pedido.Agosto;
                    sheet.Cells[i, 9].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 10].Value = Pedido.Setembro;
                    sheet.Cells[i, 10].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 11].Value = Pedido.Outubro;
                    sheet.Cells[i, 11].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 12].Value = Pedido.Novembro;
                    sheet.Cells[i, 12].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 13].Value = Pedido.Dezembro;
                    sheet.Cells[i, 13].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                    i++;
                });

                i++;

                sheet.Cells[i, 1].Value = "NF Devolução Fab";

                sheet.Cells[i, 1, i, 17].Merge = true;

                i++;

                sheet.Cells[i, 1].Value = "Nome";
                sheet.Cells[i, 2].Value = "NF";
                sheet.Cells[i, 3].Value = "Serie";
                sheet.Cells[i, 4].Value = "Emissao";
                sheet.Cells[i, 5].Value = "Janeiro";
                sheet.Cells[i, 6].Value = "Fevereiro";
                sheet.Cells[i, 7].Value = "Março";
                sheet.Cells[i, 8].Value = "Abril";
                sheet.Cells[i, 9].Value = "Maio";
                sheet.Cells[i, 10].Value = "Junho";
                sheet.Cells[i, 11].Value = "Julho";
                sheet.Cells[i, 12].Value = "Agosto";
                sheet.Cells[i, 13].Value = "Setembro";
                sheet.Cells[i, 14].Value = "Outubro";
                sheet.Cells[i, 15].Value = "Novembro";
                sheet.Cells[i, 16].Value = "Dezembro";
                sheet.Cells[i, 17].Value = "Total";

                i++;
                RelatorioDevolucao.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Nome;
                    sheet.Cells[i, 5].Value = Pedido.Janeiro;
                    sheet.Cells[i, 5].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 6].Value = Pedido.Fevereiro;
                    sheet.Cells[i, 6].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 7].Value = Pedido.Marco;
                    sheet.Cells[i, 7].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 8].Value = Pedido.Abril;
                    sheet.Cells[i, 8].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 9].Value = Pedido.Maio;
                    sheet.Cells[i, 9].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 10].Value = Pedido.Junho;
                    sheet.Cells[i, 10].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 11].Value = Pedido.Julho;
                    sheet.Cells[i, 11].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 12].Value = Pedido.Agosto;
                    sheet.Cells[i, 12].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 13].Value = Pedido.Setembro;
                    sheet.Cells[i, 13].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 14].Value = Pedido.Outubro;
                    sheet.Cells[i, 14].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 15].Value = Pedido.Novembro;
                    sheet.Cells[i, 15].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 16].Value = Pedido.Dezembro;
                    sheet.Cells[i, 16].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 17].Value = Pedido.Total;
                    sheet.Cells[i, 17].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                    sheet.Cells[i, 1, i, 4].Merge = true;
                    i++;
                    Pedido.Clientes.ForEach(cliente =>
                    {
                        sheet.Cells[i, 1].Value = cliente.Nome;

                        sheet.Cells[i, 1, i, 17].Merge = true;
                        sheet.Cells[i, 1, i, 17].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[i, 1, i, 17].Style.Fill.BackgroundColor.SetColor(1, 100, 149, 237);
                        sheet.Row(i).OutlineLevel = 1;
                        var grupo = i;
                        i++;


                        cliente.Notas.ForEach(nota =>
                        {
                            sheet.Cells[i, 1].Value = nota.Nome;
                            sheet.Cells[i, 2].Value = nota.NF;
                            sheet.Cells[i, 3].Value = nota.Serie;
                            sheet.Cells[i, 4].Value = nota.Emissao;
                            sheet.Cells[i, 5].Value = nota.Janeiro;
                            sheet.Cells[i, 5].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 6].Value = nota.Fevereiro;
                            sheet.Cells[i, 6].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 7].Value = nota.Marco;
                            sheet.Cells[i, 7].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 8].Value = nota.Abril;
                            sheet.Cells[i, 8].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 9].Value = nota.Maio;
                            sheet.Cells[i, 9].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 10].Value = nota.Junho;
                            sheet.Cells[i, 10].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 11].Value = nota.Julho;
                            sheet.Cells[i, 11].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 12].Value = nota.Agosto;
                            sheet.Cells[i, 12].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 13].Value = nota.Setembro;
                            sheet.Cells[i, 13].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 14].Value = nota.Outubro;
                            sheet.Cells[i, 14].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 15].Value = nota.Novembro;
                            sheet.Cells[i, 15].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 16].Value = nota.Dezembro;
                            sheet.Cells[i, 16].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 17].Value = nota.Total;
                            sheet.Cells[i, 17].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Row(i).OutlineLevel = 2;
                            i++;
                        });

                        
                    });
                });


                i++;

                sheet.Cells[i, 1].Value = "NF Faturamento Fab";

                sheet.Cells[i, 1, i, 17].Merge = true;

                i++;

                sheet.Cells[i, 1].Value = "Nome";
                sheet.Cells[i, 2].Value = "NF";
                sheet.Cells[i, 3].Value = "Serie";
                sheet.Cells[i, 4].Value = "Emissao";
                sheet.Cells[i, 5].Value = "Janeiro";
                sheet.Cells[i, 6].Value = "Fevereiro";
                sheet.Cells[i, 7].Value = "Março";
                sheet.Cells[i, 8].Value = "Abril";
                sheet.Cells[i, 9].Value = "Maio";
                sheet.Cells[i, 10].Value = "Junho";
                sheet.Cells[i, 11].Value = "Julho";
                sheet.Cells[i, 12].Value = "Agosto";
                sheet.Cells[i, 13].Value = "Setembro";
                sheet.Cells[i, 14].Value = "Outubro";
                sheet.Cells[i, 15].Value = "Novembro";
                sheet.Cells[i, 16].Value = "Dezembro";
                sheet.Cells[i, 17].Value = "Total";

                i++;

                RelatorioFaturamento.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Nome;
                    sheet.Cells[i, 5].Value = Pedido.Janeiro;
                    sheet.Cells[i, 5].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 6].Value = Pedido.Fevereiro;
                    sheet.Cells[i, 6].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 7].Value = Pedido.Marco;
                    sheet.Cells[i, 7].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 8].Value = Pedido.Abril;
                    sheet.Cells[i, 8].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 9].Value = Pedido.Maio;
                    sheet.Cells[i, 9].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 10].Value = Pedido.Junho;
                    sheet.Cells[i, 10].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 11].Value = Pedido.Julho;
                    sheet.Cells[i, 11].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 12].Value = Pedido.Agosto;
                    sheet.Cells[i, 12].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 13].Value = Pedido.Setembro;
                    sheet.Cells[i, 13].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 14].Value = Pedido.Outubro;
                    sheet.Cells[i, 14].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 15].Value = Pedido.Novembro;
                    sheet.Cells[i, 15].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 16].Value = Pedido.Dezembro;
                    sheet.Cells[i, 16].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                    sheet.Cells[i, 17].Value = Pedido.Total;
                    sheet.Cells[i, 17].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";

                    sheet.Cells[i, 1, i, 4].Merge = true;
                    i++;
                    Pedido.Clientes.ForEach(cliente =>
                    {
                        sheet.Cells[i, 1].Value = cliente.Nome;
                        sheet.Cells[i, 1, i, 17].Merge = true;
                        sheet.Cells[i, 1, i, 17].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[i, 1, i, 17].Style.Fill.BackgroundColor.SetColor(1, 100, 149, 237);
                        sheet.Row(i).OutlineLevel = 1;
                        i++;

                        cliente.Notas.ForEach(nota =>
                        {
                            sheet.Cells[i, 1].Value = nota.Nome;
                            sheet.Cells[i, 2].Value = nota.NF;
                            sheet.Cells[i, 3].Value = nota.Serie;
                            sheet.Cells[i, 4].Value = nota.Emissao;
                            sheet.Cells[i, 5].Value = nota.Janeiro;
                            sheet.Cells[i, 5].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 6].Value = nota.Fevereiro;
                            sheet.Cells[i, 6].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 7].Value = nota.Marco;
                            sheet.Cells[i, 7].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 8].Value = nota.Abril;
                            sheet.Cells[i, 8].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 9].Value = nota.Maio;
                            sheet.Cells[i, 9].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 10].Value = nota.Junho;
                            sheet.Cells[i, 10].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 11].Value = nota.Julho;
                            sheet.Cells[i, 11].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 12].Value = nota.Agosto;
                            sheet.Cells[i, 12].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 13].Value = nota.Setembro;
                            sheet.Cells[i, 13].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 14].Value = nota.Outubro;
                            sheet.Cells[i, 14].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 15].Value = nota.Novembro;
                            sheet.Cells[i, 15].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 16].Value = nota.Dezembro;
                            sheet.Cells[i, 16].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Cells[i, 17].Value = nota.Total;
                            sheet.Cells[i, 17].Style.Numberformat.Format = "#,##0.00;(#,##0.00)";
                            sheet.Row(i).OutlineLevel = 2;
                            i++;

                            
                        });
                    });
                });

                sheet.OutLineSummaryBelow = false;
                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FaturamentoNFFabDenuo.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FaturamentoNFFAB Excel",user);

                return LocalRedirect("/error");
            }
        }
    }
}
