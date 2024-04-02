using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;
using SGID.Models.Estoque;
using SGID.Models.Estoque.RelatorioFaturamentoNFFab;

namespace SGID.Pages.Relatorios.Estoque
{
    [Authorize]
    public class FaturamentoNFFabModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext Protheus { get; set; }

        public List<Tipo> RelatorioFaturamento { get; set; } = new List<Tipo>();

        public List<Tipo> RelatorioDevolucao { get; set; } = new List<Tipo>();

        public List<Fabricante> RelatorioFabricante { get; set; } = new List<Fabricante>();

        public FaturamentoNFFabModel(ApplicationDbContext sgid,TOTVSDENUOContext denuo)
        {
            Protheus = denuo;
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


                var CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };
                var CfNe = new int[] { 1202, 1553, 2202, 2553 };

                #region Faturado

                var queryFaturado = (from SD20 in Protheus.Sd2010s
                                     join SA10 in Protheus.Sa1010s on new { Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                     join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SC50 in Protheus.Sc5010s on new { Pedido = SD20.D2Pedido, Filial = SD20.D2Filial, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Pedido = SC50.C5Num, Filial = SC50.C5Filial, Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                                     where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" &&
                                     (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114)
                                     || ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf))
                                     && (int)(object)SD20.D2Emissao >= Inicio && (int)(object)SD20.D2Emissao <= Fim && SD20.D2Quant != 0
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
                                         TotalBrut = SD20.D2Valbrut,
                                         Fabricante = SB10.B1Fabric
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
                                         x.Emissao,
                                         x.Fabricante
                                     })
                                     .Select(x => new FaturamentoNFFAB
                                     {
                                         Filial = x.Key.Filial,
                                         Cliente = x.Key.Cliente,
                                         Loja = x.Key.Loja,
                                         Nome = x.Key.Nome,
                                         Tipo = x.Key.Tipo == "H" ? "HOSPITAL" : x.Key.Tipo == "M" ? "MEDICO" : x.Key.Tipo == "I" ? "INSTRUMENTAL" : x.Key.Tipo == "N" ? "NORMAL" : x.Key.Tipo == "C" ? "CONVENIO" : x.Key.Tipo == "P" ? "PARTICULAR" : x.Key.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                         Est = x.Key.Est,
                                         Mun = x.Key.Mun,
                                         NF = x.Key.NF,
                                         Serie = x.Key.Serie,
                                         Emissao = x.Key.Emissao,
                                         Fabricante = x.Key.Fabricante,
                                         Total = x.Sum(c => c.Total),
                                         Valipi = x.Sum(c => c.Valipi),
                                         Valicm = x.Sum(c => c.Valicm),
                                         Descon = x.Sum(c => c.Descon),
                                         Mes = x.Key.Emissao.Substring(4, 2)
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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                #endregion

                #region Devolucao

                var queryDevolucao = (from SD10 in Protheus.Sd1010s
                                      join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, NF = SD10.D1Nfori, Serie = SD10.D1Seriori, Forne = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, NF = SF20.F2Doc, Serie = SF20.F2Serie, Forne = SF20.F2Cliente, Loja = SF20.F2Loja } into Sr
                                      from A in Sr.DefaultIfEmpty()
                                      join SA10 in Protheus.Sa1010s on new { Forne = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forne = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                      where SD10.DELET != "*" && A.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*"
                                      && CfNe.Contains((int)(object)SD10.D1Cf) && (int)(object)SD10.D1Dtdigit >= Inicio && (int)(object)SD10.D1Dtdigit <= Fim
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
                                          TotalBruto = SD10.D1Total - SD10.D1Valdesc + SD10.D1Valipi,
                                          Fabricante = SB10.B1Fabric
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
                                          x.DTDIGIT,
                                          x.Fabricante
                                      }).Select(x => new FaturamentoNFFAB
                                      {
                                          Filial = x.Key.Filial,
                                          Cliente = x.Key.Cliente,
                                          Loja = x.Key.Loja,
                                          Nome = x.Key.Nome,
                                          Tipo = x.Key.Tipo == "H" ? "HOSPITAL" : x.Key.Tipo == "M" ? "MEDICO" : x.Key.Tipo == "I" ? "INSTRUMENTAL" : x.Key.Tipo == "N" ? "NORMAL" : x.Key.Tipo == "C" ? "CONVENIO" : x.Key.Tipo == "P" ? "PARTICULAR" : x.Key.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                          Est = x.Key.Est,
                                          Mun = x.Key.Mun,
                                          NF = x.Key.NF,
                                          Serie = x.Key.Serie,
                                          Emissao = x.Key.Emissao,
                                          Fabricante = x.Key.Fabricante,
                                          Total = -x.Sum(c => c.Total),
                                          Valipi = -x.Sum(c => c.Valipi),
                                          Valicm = -x.Sum(c => c.Valicm),
                                          Descon = -x.Sum(c => c.Descon),
                                          Mes = x.Key.DTDIGIT.Substring(4, 2)
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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }

                #endregion

                #region Fabricante
                var query = queryFaturado.Concat(queryDevolucao)
                    .GroupBy(x => new
                    {
                        x.Filial,
                        x.Fabricante,
                        x.Mes
                    })
                    .Select(x => new
                    {
                        x.Key.Filial,
                        x.Key.Fabricante,
                        x.Key.Mes,
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
                            switch (x.Mes)
                            {
                                case "01": NovoRelatorio.Janeiro += x.Total; break;
                                case "02": NovoRelatorio.Fevereiro += x.Total; break;
                                case "03": NovoRelatorio.Marco += x.Total; break;
                                case "04": NovoRelatorio.Abril += x.Total; break;
                                case "05": NovoRelatorio.Maio += x.Total; break;
                                case "06": NovoRelatorio.Junho += x.Total; break;
                                case "07": NovoRelatorio.Julho += x.Total; break;
                                case "08": NovoRelatorio.Agosto += x.Total; break;
                                case "09": NovoRelatorio.Setembro += x.Total; break;
                                case "10": NovoRelatorio.Outubro += x.Total; break;
                                case "11": NovoRelatorio.Novembro += x.Total; break;
                                case "12": NovoRelatorio.Dezembro += x.Total; break;
                            }

                            NovoRelatorio.Total += x.Total;
                        });

                        RelatorioFabricante.Add(NovoRelatorio);
                    }

                });

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

                var CF = new int[] { 5551, 6551, 6107, 6109, 5117, 6117 };
                var CfNe = new int[] { 1202, 1553, 2202, 2553 };


                #region Faturado

                var queryFaturado = (from SD20 in Protheus.Sd2010s
                                     join SA10 in Protheus.Sa1010s on new { Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                     join SF20 in Protheus.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                                     join SC50 in Protheus.Sc5010s on new { Pedido = SD20.D2Pedido, Filial = SD20.D2Filial, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Pedido = SC50.C5Num, Filial = SC50.C5Filial, Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli }
                                     where SD20.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" &&
                                     (((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114) || ((int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114)
                                     || ((int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114) || CF.Contains((int)(object)SD20.D2Cf))
                                     && (int)(object)SD20.D2Emissao >= Inicio && (int)(object)SD20.D2Emissao <= Fim && SD20.D2Quant != 0
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
                                         TotalBrut = SD20.D2Valbrut,
                                         Fabricante = SB10.B1Fabric
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
                                         x.Emissao,
                                         x.Fabricante
                                     })
                                     .Select(x => new FaturamentoNFFAB
                                     {
                                         Filial = x.Key.Filial,
                                         Cliente = x.Key.Cliente,
                                         Loja = x.Key.Loja,
                                         Nome = x.Key.Nome,
                                         Tipo = x.Key.Tipo == "H" ? "HOSPITAL" : x.Key.Tipo == "M" ? "MEDICO" : x.Key.Tipo == "I" ? "INSTRUMENTAL" : x.Key.Tipo == "N" ? "NORMAL" : x.Key.Tipo == "C" ? "CONVENIO" : x.Key.Tipo == "P" ? "PARTICULAR" : x.Key.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                         Est = x.Key.Est,
                                         Mun = x.Key.Mun,
                                         NF = x.Key.NF,
                                         Serie = x.Key.Serie,
                                         Emissao = x.Key.Emissao,
                                         Fabricante = x.Key.Fabricante,
                                         Total = x.Sum(c => c.Total),
                                         Valipi = x.Sum(c => c.Valipi),
                                         Valicm = x.Sum(c => c.Valicm),
                                         Descon = x.Sum(c => c.Descon),
                                         Mes = x.Key.Emissao.Substring(4, 2)
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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioFaturamento.Add(tipo);
                }
                #endregion

                #region Devolucao

                var queryDevolucao = (from SD10 in Protheus.Sd1010s
                                      join SF20 in Protheus.Sf2010s on new { Filial = SD10.D1Filial, NF = SD10.D1Nfori, Serie = SD10.D1Seriori, Forne = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, NF = SF20.F2Doc, Serie = SF20.F2Serie, Forne = SF20.F2Cliente, Loja = SF20.F2Loja } into Sr
                                      from A in Sr.DefaultIfEmpty()
                                      join SA10 in Protheus.Sa1010s on new { Forne = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forne = SA10.A1Cod, Loja = SA10.A1Loja }
                                      join SB10 in Protheus.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                                      where SD10.DELET != "*" && A.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*"
                                      && CfNe.Contains((int)(object)SD10.D1Cf) && (int)(object)SD10.D1Dtdigit >= Inicio && (int)(object)SD10.D1Dtdigit <= Fim
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
                                          TotalBruto = SD10.D1Total - SD10.D1Valdesc + SD10.D1Valipi,
                                          Fabricante = SB10.B1Fabric
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
                                          x.DTDIGIT,
                                          x.Fabricante
                                      }).Select(x => new FaturamentoNFFAB
                                      {
                                          Filial = x.Key.Filial,
                                          Cliente = x.Key.Cliente,
                                          Loja = x.Key.Loja,
                                          Nome = x.Key.Nome,
                                          Tipo = x.Key.Tipo == "H" ? "HOSPITAL" : x.Key.Tipo == "M" ? "MEDICO" : x.Key.Tipo == "I" ? "INSTRUMENTAL" : x.Key.Tipo == "N" ? "NORMAL" : x.Key.Tipo == "C" ? "CONVENIO" : x.Key.Tipo == "P" ? "PARTICULAR" : x.Key.Tipo == "S" ? "SUB-DISTRIBUIDOR" : "OUTROS",
                                          Est = x.Key.Est,
                                          Mun = x.Key.Mun,
                                          NF = x.Key.NF,
                                          Serie = x.Key.Serie,
                                          Emissao = x.Key.Emissao,
                                          Fabricante = x.Key.Fabricante,
                                          Total = -x.Sum(c => c.Total),
                                          Valipi = -x.Sum(c => c.Valipi),
                                          Valicm = -x.Sum(c => c.Valicm),
                                          Descon = -x.Sum(c => c.Descon),
                                          Mes = x.Key.DTDIGIT.Substring(4, 2)
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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

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
                                    NF = nota.NF,
                                    Emissao = nota.Emissao
                                };

                                switch (nota.Mes)
                                {
                                    case "01":
                                        {
                                            tipo.Janeiro += nota.Total;
                                            NotaFiscal.Janeiro += nota.Total; break;
                                        }
                                    case "02":
                                        {
                                            tipo.Fevereiro += nota.Total;
                                            NotaFiscal.Fevereiro += nota.Total; break;
                                        }
                                    case "03":
                                        {
                                            tipo.Marco += nota.Total;
                                            NotaFiscal.Marco += nota.Total; break;
                                        }
                                    case "04":
                                        {
                                            tipo.Abril += nota.Total;
                                            NotaFiscal.Abril += nota.Total; break;
                                        }
                                    case "05":
                                        {
                                            tipo.Maio += nota.Total;
                                            NotaFiscal.Maio += nota.Total; break;
                                        }
                                    case "06":
                                        {
                                            tipo.Junho += nota.Total;
                                            NotaFiscal.Junho += nota.Total; break;
                                        }
                                    case "07":
                                        {
                                            tipo.Julho += nota.Total;
                                            NotaFiscal.Julho += nota.Total; break;
                                        }
                                    case "08":
                                        {
                                            tipo.Agosto += nota.Total;
                                            NotaFiscal.Agosto += nota.Total; break;
                                        }
                                    case "09":
                                        {
                                            tipo.Setembro += nota.Total;
                                            NotaFiscal.Setembro += nota.Total; break;
                                        }
                                    case "10":
                                        {
                                            tipo.Outubro += nota.Total;
                                            NotaFiscal.Outubro += nota.Total; break;
                                        }
                                    case "11":
                                        {
                                            tipo.Novembro += nota.Total;
                                            NotaFiscal.Novembro += nota.Total; break;
                                        }
                                    case "12":
                                        {
                                            tipo.Dezembro += nota.Total;
                                            NotaFiscal.Dezembro += nota.Total; break;
                                        }
                                }

                                tipo.Total += x.Total;
                                NotaFiscal.Total += x.Total;

                                cliente.Notas.Add(NotaFiscal);

                            });

                            tipo.Clientes.Add(cliente);
                        }
                    });

                    RelatorioDevolucao.Add(tipo);
                }

                #endregion

                #region Fabricante
                var query = queryFaturado.Concat(queryDevolucao)
                    .GroupBy(x => new
                    {
                        x.Filial,
                        x.Fabricante,
                        x.Mes
                    })
                    .Select(x => new
                    {
                        x.Key.Filial,
                        x.Key.Fabricante,
                        x.Key.Mes,
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
                            switch (x.Mes)
                            {
                                case "01": NovoRelatorio.Janeiro += x.Total; break;
                                case "02": NovoRelatorio.Fevereiro += x.Total; break;
                                case "03": NovoRelatorio.Marco += x.Total; break;
                                case "04": NovoRelatorio.Abril += x.Total; break;
                                case "05": NovoRelatorio.Maio += x.Total; break;
                                case "06": NovoRelatorio.Junho += x.Total; break;
                                case "07": NovoRelatorio.Julho += x.Total; break;
                                case "08": NovoRelatorio.Agosto += x.Total; break;
                                case "09": NovoRelatorio.Setembro += x.Total; break;
                                case "10": NovoRelatorio.Outubro += x.Total; break;
                                case "11": NovoRelatorio.Novembro += x.Total; break;
                                case "12": NovoRelatorio.Dezembro += x.Total; break;
                            }

                            NovoRelatorio.Total += x.Total;
                        });

                        RelatorioFabricante.Add(NovoRelatorio);
                    }

                });

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
