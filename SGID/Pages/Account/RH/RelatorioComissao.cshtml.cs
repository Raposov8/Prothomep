using DocumentFormat.OpenXml.Vml.Office;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Models;
using SGID.Models.Account.RH;
using SGID.Models.Controladoria.FaturamentoNF;
using SGID.Models.Denuo;
using SGID.Models.Estoque.RelatorioFaturamentoNFFab;
using SGID.Models.Inter;
using System.Linq;

namespace SGID.Pages.Account.RH
{
    public class RelatorioComissaoModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public string MesAno { get; set; }
        public string Ano { get; set; }
        public int IdTime { get; set; }

        public string Mes { get; set; }

        public TimeRH Representante { get; set; }

        public List<RelatorioFaturamentoTime> Faturamento { get; set; } = new List<RelatorioFaturamentoTime>();
        public List<RelatorioFaturamentoTime> GestorFaturamento { get; set; } = new List<RelatorioFaturamentoTime>();
        public List<RelatorioFaturamentoTime> LinhaFaturamento { get; set; } = new List<RelatorioFaturamentoTime>();

        public List<RelatorioDevolucaoFat> Devolucao { get; set; } = new List<RelatorioDevolucaoFat>();
        public List<RelatorioDevolucaoFat> GestorDevolucao { get; set; } = new List<RelatorioDevolucaoFat>();
        public List<RelatorioDevolucaoFat> LinhaDevolucao { get; set; } = new List<RelatorioDevolucaoFat>();

        public RelatorioComissaoModel(ApplicationDbContext sgid,TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            SGID = sgid;
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
        }

        public void OnGet(string Mes,string Ano,int IdTime)
        {
            this.Ano = Ano;
            MesAno = Mes;
            this.IdTime = IdTime;

            string Tempo = $"{Mes}/01/{Ano}";

            var date = DateTime.Parse(Tempo);

            string data = date.ToString("yyyy/MM").Replace("/", "");
            this.Mes = date.ToString("MMMM").ToUpper();
            string DataInicio = data + "01";
            string DataFim = data + "31";
            int[] CF = new int[] { 5551, 6551, 6107, 6109 };

            #region INTERMEDIC

            #region Faturado
            var query = (from SD20 in ProtheusInter.Sd2010s
                         join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                         join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                         join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                         join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                         && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
                         && (int)(object)SC50.C5XDtcir >= 20200801
                         select new 
                         {
                             NF = SD20.D2Doc,
                             Cliente = SA10.A1Nome,
                             Total = SD20.D2Total,
                             Login = SA30.A3Xlogin,
                             Gestor = SA30.A3Xlogsup,
                             Linha = SA30.A3Xdescun
                         });

            var resultadoInter = query.GroupBy(x => new
            {
                x.NF,
                x.Cliente,
                x.Total,
                x.Login,
                x.Gestor,
                x.Linha
            }).Select(x => new RelatorioFaturamentoTime
            {
                NF = x.Key.NF,
                Cliente = x.Key.Cliente,
                Total = x.Sum(c => c.Total),
                Login = x.Key.Login.Trim(),
                Gestor = x.Key.Gestor.Trim(),
                Linha = x.Key.Linha.Trim(),
                Empresa = "INTERMEDIC"
            }).ToList();
            #endregion

            #region Devolucao

            var DevolucaoInter = new List<RelatorioDevolucaoFat>();

            var teste = (from SD10 in ProtheusInter.Sd1010s
                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                         join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                         join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                         join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                         where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                         && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                         && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                         && (int)(object)SD10.D1Dtdigit >= 20200801
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
                             SA30.A3Xlogin,
                             SA30.A3Xlogsup,
                             Linha = SA30.A3Xdescun
                         }
                         ).ToList();

            if (teste.Count != 0)
            {
                teste.ForEach(x =>
                {
                    if (!DevolucaoInter.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                    {

                        var Iguais = teste
                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                        double desconto = 0;
                        double valipi = 0;
                        double total = 0;
                        double valicm = 0;
                        Iguais.ForEach(x =>
                        {
                            desconto += x.D1Valdesc;
                            valipi += x.D1Valipi;
                            total += x.D1Total;
                            valicm += x.D1Valicm;
                        });

                        DevolucaoInter.Add(new RelatorioDevolucaoFat
                        {
                            Filial = x.D1Filial,
                            Clifor = x.D1Fornece,
                            Loja = x.D1Loja,
                            Nome = x.A1Nome,
                            Tipo = x.A1Clinter,
                            Nf = x.D1Doc,
                            Serie = x.D1Serie,
                            Digitacao = x.D1Dtdigit,
                            Total = total - desconto,
                            Valipi = valipi,
                            Valicm = valicm,
                            Descon = desconto,
                            TotalBrut = total - desconto + valipi,
                            A3Nome = x.A3Nome,
                            D1Nfori = x.D1Nfori,
                            D1Seriori = x.D1Seriori,
                            D1Datori = x.D1Datori,
                            Login = x.A3Xlogin.Trim(),
                            Gestor = x.A3Xlogsup.Trim(),
                            Linha = x.Linha,
                            Empresa = "INTERMEDIC"
                        });
                    }
                });
            }

            #endregion

            #endregion

            #region DENUO

            #region Faturado

            var query2 = (from SD20 in ProtheusDenuo.Sd2010s
                          join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                          join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                          join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                          join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                          join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                          where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                          && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                          (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                          && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
                          && (int)(object)SC50.C5XDtcir >= 20200801
                          select new
                          {
                              NF = SD20.D2Doc,
                              Total = SD20.D2Total,
                              Login = SA30.A3Xlogin,
                              Gestor = SA30.A3Xlogsup,
                              Linha = SA30.A3Xdescun
                          });

            var resultadoDenuo = query2.GroupBy(x => new
            {
                x.NF,
                x.Total,
                x.Login,
                x.Gestor,
                x.Linha
            }).Select(x => new RelatorioFaturamentoTime
            {
                NF = x.Key.NF,
                Total = x.Sum(c => c.Total),
                Login = x.Key.Login.Trim(),
                Gestor = x.Key.Gestor.Trim(),
                Linha = x.Key.Linha.Trim()
            }).ToList();
            #endregion

            #region Devolucao
            var DevolucaoDenuo = new List<RelatorioDevolucaoFat>();

            var teste2 = (from SD10 in ProtheusDenuo.Sd1010s
                          join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                          join SA30 in ProtheusDenuo.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                          join SA10 in ProtheusDenuo.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                          join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                          where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                          && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                          && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                          && (int)(object)SD10.D1Dtdigit >= 20200801
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
                              SA30.A3Xlogin,
                              SA30.A3Xlogsup,
                              SD10.D1Nfori,
                              SD10.D1Seriori,
                              SD10.D1Datori,
                              SD10.D1Emissao,
                              Linha = SA30.A3Xdescun
                          }
                         ).ToList();

            if (teste2.Count != 0)
            {
                teste2.ForEach(x =>
                {
                    if (!DevolucaoDenuo.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                    {

                        var Iguais = teste2
                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                        double desconto = 0;
                        double valipi = 0;
                        double total = 0;
                        double valicm = 0;
                        Iguais.ForEach(x =>
                        {
                            desconto += x.D1Valdesc;
                            valipi += x.D1Valipi;
                            total += x.D1Total;
                            valicm += x.D1Valicm;
                        });

                        DevolucaoDenuo.Add(new RelatorioDevolucaoFat
                        {
                            Filial = x.D1Filial,
                            Clifor = x.D1Fornece,
                            Loja = x.D1Loja,
                            Nome = x.A1Nome,
                            Tipo = x.A1Clinter,
                            Nf = x.D1Doc,
                            Serie = x.D1Serie,
                            Digitacao = x.D1Dtdigit,
                            Total = total - desconto,
                            Valipi = valipi,
                            Valicm = valicm,
                            Descon = desconto,
                            TotalBrut = total - desconto + valipi,
                            A3Nome = x.A3Nome,
                            D1Nfori = x.D1Nfori,
                            D1Seriori = x.D1Seriori,
                            D1Datori = x.D1Datori,
                            Login = x.A3Xlogin.Trim(),
                            Gestor = x.A3Xlogsup.Trim(),
                            Linha = x.Linha
                        });
                    }
                });
            }
            #endregion

            #endregion

            var usuario = SGID.Times.FirstOrDefault(x=> x.Id == IdTime);

            var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == usuario.Id).ToList();

            Representante = new TimeRH
            {
                User = usuario
            };

            Representante.Faturado += resultadoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total);

            Representante.Faturado += resultadoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total);

            Faturamento.AddRange(resultadoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
            Faturamento.AddRange(resultadoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
            Devolucao.AddRange(DevolucaoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
            Devolucao.AddRange(DevolucaoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());

            if (usuario.Integrante.ToUpper() == "MICHEL.SAMPAIO")
            {
                Representante.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total);
            }
            else
            {
                Representante.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper()).Sum(x => x.Total);
            }

            if (usuario.Integrante.ToUpper() == "MICHEL.SAMPAIO")
            {
                Representante.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES").Sum(x => x.Total);
                GestorFaturamento.AddRange(resultadoInter.Where(x => x.Gestor == "ANDRE.SALES").ToList());
                GestorFaturamento.AddRange(resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES").ToList());
                GestorDevolucao.AddRange(DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES").ToList());
                GestorDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES").ToList());
            }
            else
            {
                Representante.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper().ToUpper()).Sum(x => x.Total);
                GestorFaturamento.AddRange(resultadoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                GestorFaturamento.AddRange(resultadoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                GestorDevolucao.AddRange(DevolucaoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                GestorDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
            }


            Produtos.ForEach(prod =>
            {
                Representante.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                if (usuario.Integrante.ToUpper() == "TIAGO.FONSECA")
                {
                    Representante.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO" ) || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && (x.Gestor != "RONAN.JOVINO" ) || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total);
                }
                else
                {
                    Representante.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" ).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" ).Sum(x => x.Total);
                }

                LinhaFaturamento.AddRange(resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").ToList());
                LinhaDevolucao.AddRange(DevolucaoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").ToList());

                if (usuario.Integrante.ToUpper() == "TIAGO.FONSECA")
                {
                    LinhaFaturamento.AddRange(resultadoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").ToList());
                    LinhaDevolucao.AddRange(DevolucaoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" ) || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").ToList());
                }
                else
                {
                    LinhaFaturamento.AddRange(resultadoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                    LinhaDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                }
            });


            Representante.Comissao = Representante.Faturado * (Representante.User.Porcentagem / 100);
            Representante.ComissaoEquipe = Representante.FaturadoEquipe * (Representante.User.PorcentagemSeg / 100);
            Representante.ComissaoProduto = Representante.FaturadoProduto * (Representante.User.PorcentagemGenProd / 100);

        }

        public IActionResult OnPostExport(string Mes, string Ano, int IdTime)
        {

            string Tempo = $"{Mes}/01/{Ano}";

            var date = DateTime.Parse(Tempo);

            string data = date.ToString("yyyy/MM").Replace("/", "");
            this.Mes = date.ToString("MMMM").ToUpper();
            string DataInicio = data + "01";
            string DataFim = data + "31";
            int[] CF = new int[] { 5551, 6551, 6107, 6109 };

            #region INTERMEDIC

            #region Faturado
            var query = (from SD20 in ProtheusInter.Sd2010s
                         join SA10 in ProtheusInter.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                         join SB10 in ProtheusInter.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                         join SC50 in ProtheusInter.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                         join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                         && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                         (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                         && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
                         && (int)(object)SC50.C5XDtcir >= 20200801
                         select new
                         {
                             NF = SD20.D2Doc,
                             Cliente = SA10.A1Nome,
                             Total = SD20.D2Total,
                             Login = SA30.A3Xlogin,
                             Gestor = SA30.A3Xlogsup,
                             Linha = SA30.A3Xdescun,
                             Dor = SB10.B1Ugrpint
                         });

            var resultadoInter = query.GroupBy(x => new
            {
                x.NF,
                x.Cliente,
                x.Total,
                x.Login,
                x.Gestor,
                x.Linha,
                x.Dor
            }).Select(x => new RelatorioFaturamentoTime
            {
                NF = x.Key.NF,
                Cliente = x.Key.Cliente,
                Total = x.Sum(c => c.Total),
                Login = x.Key.Login.Trim(),
                Gestor = x.Key.Gestor.Trim(),
                Linha = x.Key.Linha.Trim(),
                Empresa = "INTERMEDIC",
                Dor = x.Key.Dor
            }).ToList();
            #endregion

            #region Devolucao

            var DevolucaoInter = new List<RelatorioDevolucaoFat>();

            var teste = (from SD10 in ProtheusInter.Sd1010s
                         join SF20 in ProtheusInter.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                         join SA30 in ProtheusInter.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                         join SA10 in ProtheusInter.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                         join SB10 in ProtheusInter.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                         where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                         && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                         && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                         && (int)(object)SD10.D1Dtdigit >= 20200801
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
                             SA30.A3Xlogin,
                             SA30.A3Xlogsup,
                             Linha = SA30.A3Xdescun,
                             Dor = SB10.B1Ugrpint
                         }
                         ).ToList();

            if (teste.Count != 0)
            {
                teste.ForEach(x =>
                {
                    if (!DevolucaoInter.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                    {

                        var Iguais = teste
                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                        double desconto = 0;
                        double valipi = 0;
                        double total = 0;
                        double valicm = 0;
                        Iguais.ForEach(x =>
                        {
                            desconto += x.D1Valdesc;
                            valipi += x.D1Valipi;
                            total += x.D1Total;
                            valicm += x.D1Valicm;
                        });

                        DevolucaoInter.Add(new RelatorioDevolucaoFat
                        {
                            Filial = x.D1Filial,
                            Clifor = x.D1Fornece,
                            Loja = x.D1Loja,
                            Nome = x.A1Nome,
                            Tipo = x.A1Clinter,
                            Nf = x.D1Doc,
                            Serie = x.D1Serie,
                            Digitacao = x.D1Dtdigit,
                            Total = total - desconto,
                            Valipi = valipi,
                            Valicm = valicm,
                            Descon = desconto,
                            TotalBrut = total - desconto + valipi,
                            A3Nome = x.A3Nome,
                            D1Nfori = x.D1Nfori,
                            D1Seriori = x.D1Seriori,
                            D1Datori = x.D1Datori,
                            Login = x.A3Xlogin.Trim(),
                            Gestor = x.A3Xlogsup.Trim(),
                            Linha = x.Linha,
                            Empresa = "INTERMEDIC",
                            DOR = x.Dor
                        });
                    }
                });
            }

            #endregion

            #endregion

            #region DENUO

            #region Faturado

            var query2 = (from SD20 in ProtheusDenuo.Sd2010s
                          join SA10 in ProtheusDenuo.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                          join SB10 in ProtheusDenuo.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                          join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD20.D2Filial, Doc = SD20.D2Doc, Serie = SD20.D2Serie, Cliente = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Cliente = SF20.F2Cliente, Loja = SF20.F2Loja }
                          join SC50 in ProtheusDenuo.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                          join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                          where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SF20.DELET != "*" && SC50.DELET != "*" && SA30.DELET != "*"
                          && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114 ||
                          (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || CF.Contains((int)(object)SD20.D2Cf)) && ((int)(object)SD20.D2Emissao >= (int)(object)DataInicio && (int)(object)SD20.D2Emissao <= (int)(object)DataFim)
                          && SD20.D2Quant != 0 && SC50.C5Utpoper == "F" && SA10.A1Clinter != "S" && SA10.A1Cgc != "04715053000140" && SA10.A1Cgc != "04715053000220" && SA10.A1Cgc != "01390500000140" && (int)(object)SD20.D2Emissao >= 20200801
                          && (int)(object)SC50.C5XDtcir >= 20200801
                          select new
                          {
                              NF = SD20.D2Doc,
                              Total = SD20.D2Total,
                              Login = SA30.A3Xlogin,
                              Gestor = SA30.A3Xlogsup,
                              Linha = SA30.A3Xdescun
                          });

            var resultadoDenuo = query2.GroupBy(x => new
            {
                x.NF,
                x.Total,
                x.Login,
                x.Gestor,
                x.Linha
            }).Select(x => new RelatorioFaturamentoTime
            {
                NF = x.Key.NF,
                Total = x.Sum(c => c.Total),
                Login = x.Key.Login.Trim(),
                Gestor = x.Key.Gestor.Trim(),
                Linha = x.Key.Linha.Trim()
            }).ToList();
            #endregion

            #region Devolucao
            var DevolucaoDenuo = new List<RelatorioDevolucaoFat>();

            var teste2 = (from SD10 in ProtheusDenuo.Sd1010s
                          join SF20 in ProtheusDenuo.Sf2010s on new { Filial = SD10.D1Filial, Doc = SD10.D1Nfori, Serie = SD10.D1Seriori, Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Filial = SF20.F2Filial, Doc = SF20.F2Doc, Serie = SF20.F2Serie, Forn = SF20.F2Cliente, Loja = SF20.F2Loja }
                          join SA30 in ProtheusDenuo.Sa3010s on SF20.F2Vend1 equals SA30.A3Cod
                          join SA10 in ProtheusDenuo.Sa1010s on new { Forn = SD10.D1Fornece, Loja = SD10.D1Loja } equals new { Forn = SA10.A1Cod, Loja = SA10.A1Loja }
                          join SB10 in ProtheusDenuo.Sb1010s on SD10.D1Cod equals SB10.B1Cod
                          where SD10.DELET != "*" && SF20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SA30.DELET != "*"
                          && SA30.A3Xunnego != "000008" && (SD10.D1Cf == "1202" || SD10.D1Cf == "2202" || SD10.D1Cf == "3202" || SD10.D1Cf == "1553" || SD10.D1Cf == "2553")
                          && (int)(object)SD10.D1Dtdigit >= (int)(object)DataInicio && (int)(object)SD10.D1Dtdigit <= (int)(object)DataFim
                          && (int)(object)SD10.D1Dtdigit >= 20200801
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
                              SA30.A3Xlogin,
                              SA30.A3Xlogsup,
                              SD10.D1Nfori,
                              SD10.D1Seriori,
                              SD10.D1Datori,
                              SD10.D1Emissao,
                              Linha = SA30.A3Xdescun
                          }
                         ).ToList();

            if (teste2.Count != 0)
            {
                teste2.ForEach(x =>
                {
                    if (!DevolucaoDenuo.Any(d => d.Nome == x.A1Nome && d.Nf == x.D1Doc))
                    {

                        var Iguais = teste2
                        .Where(c => c.A1Nome == x.A1Nome && c.A3Nome == x.A3Nome && c.D1Filial == x.D1Filial
                        && c.D1Fornece == x.D1Fornece && c.D1Loja == x.D1Loja && c.A1Clinter == x.A1Clinter
                        && c.D1Doc == x.D1Doc && c.D1Serie == x.D1Serie && c.D1Emissao == x.D1Emissao && c.D1Dtdigit == x.D1Dtdigit
                        && c.D1Nfori == x.D1Nfori && c.D1Seriori == x.D1Seriori && c.D1Datori == x.D1Datori).ToList();

                        double desconto = 0;
                        double valipi = 0;
                        double total = 0;
                        double valicm = 0;
                        Iguais.ForEach(x =>
                        {
                            desconto += x.D1Valdesc;
                            valipi += x.D1Valipi;
                            total += x.D1Total;
                            valicm += x.D1Valicm;
                        });

                        DevolucaoDenuo.Add(new RelatorioDevolucaoFat
                        {
                            Filial = x.D1Filial,
                            Clifor = x.D1Fornece,
                            Loja = x.D1Loja,
                            Nome = x.A1Nome,
                            Tipo = x.A1Clinter,
                            Nf = x.D1Doc,
                            Serie = x.D1Serie,
                            Digitacao = x.D1Dtdigit,
                            Total = total - desconto,
                            Valipi = valipi,
                            Valicm = valicm,
                            Descon = desconto,
                            TotalBrut = total - desconto + valipi,
                            A3Nome = x.A3Nome,
                            D1Nfori = x.D1Nfori,
                            D1Seriori = x.D1Seriori,
                            D1Datori = x.D1Datori,
                            Login = x.A3Xlogin.Trim(),
                            Gestor = x.A3Xlogsup.Trim(),
                            Linha = x.Linha
                        });
                    }
                });
            }
            #endregion

            #endregion

            var usuario = SGID.Times.FirstOrDefault(x => x.Id == IdTime);

            var Produtos = SGID.TimeProdutos.Where(c => c.TimeId == usuario.Id).ToList();

            Representante = new TimeRH
            {
                User = usuario
            };

            Representante.Faturado += resultadoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total);

            Representante.Faturado += resultadoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).Sum(x => x.Total);

            Faturamento.AddRange(resultadoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
            Faturamento.AddRange(resultadoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
            Devolucao.AddRange(DevolucaoInter.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());
            Devolucao.AddRange(DevolucaoDenuo.Where(x => x.Login == usuario.Integrante.ToUpper()).ToList());

            if (usuario.Integrante.ToUpper() == "MICHEL.SAMPAIO")
            {
                Representante.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.Dor != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                Representante.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.Dor != "082").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").Sum(x => x.Total);
                GestorFaturamento.AddRange(resultadoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.Dor != "082").ToList());
                GestorFaturamento.AddRange(resultadoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.Dor != "082").ToList());
                GestorDevolucao.AddRange(DevolucaoInter.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").ToList());
                GestorDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Gestor == "ANDRE.SALES" && x.DOR != "082").ToList());
            }
            else
            {
                Representante.FaturadoEquipe += resultadoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper() && x.Dor != "082").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper() && x.DOR != "082").Sum(x => x.Total);
                Representante.FaturadoEquipe += resultadoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper().ToUpper()).Sum(x => x.Total);
                GestorFaturamento.AddRange(resultadoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                GestorFaturamento.AddRange(resultadoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                GestorDevolucao.AddRange(DevolucaoInter.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
                GestorDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Gestor == usuario.Integrante.ToUpper()).ToList());
            }


            Produtos.ForEach(prod =>
            {
                int i = 0;
                if (usuario.Integrante.ToUpper() == "TIAGO.FONSECA")
                {
                    if (i == 0)
                    {
                        Representante.FaturadoProduto += resultadoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total) - DevolucaoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total);
                        Representante.FaturadoProduto += resultadoInter.Where(x => (x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total) - DevolucaoInter.Where(x => (x.Linha.Trim() == prod.Produto && x.Gestor != "RONAN.JOVINO")|| x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").Sum(x => x.Total);
                        LinhaFaturamento.AddRange(resultadoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").ToList());
                        LinhaDevolucao.AddRange(DevolucaoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO") || x.Login == "RICARDO.RAMOS" || x.Login == "JULIANO.SOARES" || x.Login == "ELAINE.MARTINS").ToList());
                        i++;
                    }
                    else
                    {
                        Representante.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                        Representante.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.Gestor != "RONAN.JOVINO").Sum(x => x.Total);
                        LinhaFaturamento.AddRange(resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").ToList());
                        LinhaDevolucao.AddRange(DevolucaoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO").ToList());
                        LinhaFaturamento.AddRange(resultadoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO")).ToList());
                        LinhaDevolucao.AddRange(DevolucaoDenuo.Where(x => (x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO")).ToList());
                    }
                }
                else
                {
                    Representante.FaturadoProduto += resultadoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoDenuo.Where(x => x.Linha.Trim() == prod.Produto.Trim() && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                    Representante.FaturadoProduto += resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total) - DevolucaoInter.Where(x => x.Linha.Trim() == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").Sum(x => x.Total);
                    LinhaFaturamento.AddRange(resultadoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                    LinhaDevolucao.AddRange(DevolucaoInter.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                    LinhaFaturamento.AddRange(resultadoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                    LinhaDevolucao.AddRange(DevolucaoDenuo.Where(x => x.Linha == prod.Produto && x.Gestor != "RONAN.JOVINO" && x.Gestor != "TIAGO.FONSECA").ToList());
                }
            });


            Representante.Comissao = Representante.Faturado * (Representante.User.Porcentagem / 100);
            Representante.ComissaoEquipe = Representante.FaturadoEquipe * (Representante.User.PorcentagemSeg / 100);
            Representante.ComissaoProduto = Representante.FaturadoProduto * (Representante.User.PorcentagemGenProd / 100);

            using ExcelPackage package = new ExcelPackage();
            package.Workbook.Worksheets.Add("RelatorioComissaoRH");

            var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "RelatorioComissaoRH");

            sheet.Cells[1, 1].Value = "Faturamento Venda Direta";
            sheet.Cells[1, 1, 1, 7].Merge = true;

            sheet.Cells[2, 1].Value = "Empresa";
            sheet.Cells[2, 2].Value = "NF";
            sheet.Cells[2, 3].Value = "CLIENTE";
            sheet.Cells[2, 4].Value = "VALOR";
            sheet.Cells[2, 5].Value = "VENDEDOR";
            sheet.Cells[2, 6].Value = "GESTOR";
            sheet.Cells[2, 7].Value = "LINHA";

            int i = 3;

            Faturamento.ForEach(Pedido =>
            {
                sheet.Cells[i, 1].Value = Pedido.Empresa;
                sheet.Cells[i, 2].Value = Pedido.NF;
                sheet.Cells[i, 3].Value = Pedido.Cliente;
                sheet.Cells[i, 4].Value = Pedido.Total;
                sheet.Cells[i, 5].Value = Pedido.Login;
                sheet.Cells[i, 6].Value = Pedido.Gestor;
                sheet.Cells[i, 7].Value = Pedido.Linha;

                i++;
            });

            i++;

            sheet.Cells[i, 1].Value = "Devolução Venda Direta";
            sheet.Cells[i, 1, i, 7].Merge = true;
            i++;

            sheet.Cells[i, 1].Value = "Empresa";
            sheet.Cells[i, 2].Value = "NF";
            sheet.Cells[i, 3].Value = "CLIENTE";
            sheet.Cells[i, 4].Value = "VALOR";
            sheet.Cells[i, 5].Value = "VENDEDOR";
            sheet.Cells[i, 6].Value = "GESTOR";
            sheet.Cells[i, 7].Value = "LINHA";
            i++;

            Devolucao.ForEach(Pedido =>
            {
                sheet.Cells[i, 1].Value = Pedido.Empresa;
                sheet.Cells[i, 2].Value = Pedido.D1Nfori;
                sheet.Cells[i, 3].Value = Pedido.Nome;
                sheet.Cells[i, 4].Value = Pedido.Total;
                sheet.Cells[i, 5].Value = Pedido.Login;
                sheet.Cells[i, 6].Value = Pedido.Gestor;
                sheet.Cells[i, 7].Value = Pedido.Linha;

                i++;
            });

            i++;

            sheet.Cells[i, 1].Value = "Faturamento Sobre Equipe";
            sheet.Cells[i, 1, i, 7].Merge = true;

            i++;

            sheet.Cells[i, 1].Value = "Empresa";
            sheet.Cells[i, 2].Value = "NF";
            sheet.Cells[i, 3].Value = "CLIENTE";
            sheet.Cells[i, 4].Value = "VALOR";
            sheet.Cells[i, 5].Value = "VENDEDOR";
            sheet.Cells[i, 6].Value = "GESTOR";
            sheet.Cells[i, 7].Value = "LINHA";
            i++;

            GestorFaturamento.ForEach(Pedido =>
            {
                sheet.Cells[i, 1].Value = Pedido.Empresa;
                sheet.Cells[i, 2].Value = Pedido.NF;
                sheet.Cells[i, 3].Value = Pedido.Cliente;
                sheet.Cells[i, 4].Value = Pedido.Total;
                sheet.Cells[i, 5].Value = Pedido.Login;
                sheet.Cells[i, 6].Value = Pedido.Gestor;
                sheet.Cells[i, 7].Value = Pedido.Linha;

                i++;
            });

            i++;

            sheet.Cells[i, 1].Value = "Devolução Sobre Equipe";
            sheet.Cells[i, 1, i, 7].Merge = true;

            i++;

            sheet.Cells[i, 1].Value = "Empresa";
            sheet.Cells[i, 2].Value = "NF";
            sheet.Cells[i, 3].Value = "CLIENTE";
            sheet.Cells[i, 4].Value = "VALOR";
            sheet.Cells[i, 5].Value = "VENDEDOR";
            sheet.Cells[i, 6].Value = "GESTOR";
            sheet.Cells[i, 7].Value = "LINHA";
            i++;

            GestorDevolucao.ForEach(Pedido =>
            {
                sheet.Cells[i, 1].Value = Pedido.Empresa;
                sheet.Cells[i, 2].Value = Pedido.D1Nfori;
                sheet.Cells[i, 3].Value = Pedido.Nome;
                sheet.Cells[i, 4].Value = Pedido.Total;
                sheet.Cells[i, 5].Value = Pedido.Login;
                sheet.Cells[i, 6].Value = Pedido.Gestor;
                sheet.Cells[i, 7].Value = Pedido.Linha;

                i++;
            });

            i++;

            sheet.Cells[i, 1].Value = "Faturamento Sobre Linha";
            sheet.Cells[i, 1, i, 7].Merge = true;

            i++;

            sheet.Cells[i, 1].Value = "Empresa";
            sheet.Cells[i, 2].Value = "NF";
            sheet.Cells[i, 3].Value = "CLIENTE";
            sheet.Cells[i, 4].Value = "VALOR";
            sheet.Cells[i, 5].Value = "VENDEDOR";
            sheet.Cells[i, 6].Value = "GESTOR";
            sheet.Cells[i, 7].Value = "LINHA";

            i++;

            LinhaFaturamento.ForEach(Pedido =>
            {
                sheet.Cells[i, 1].Value = Pedido.Empresa;
                sheet.Cells[i, 2].Value = Pedido.NF;
                sheet.Cells[i, 3].Value = Pedido.Cliente;
                sheet.Cells[i, 4].Value = Pedido.Total;
                sheet.Cells[i, 5].Value = Pedido.Login;
                sheet.Cells[i, 6].Value = Pedido.Gestor;
                sheet.Cells[i, 7].Value = Pedido.Linha;

                i++;
            });

            i++;

            sheet.Cells[i, 1].Value = "Devolução Sobre Linha";
            sheet.Cells[i, 1, i, 7].Merge = true;

            i++;

            sheet.Cells[i, 1].Value = "Empresa";
            sheet.Cells[i, 2].Value = "NF";
            sheet.Cells[i, 3].Value = "CLIENTE";
            sheet.Cells[i, 4].Value = "VALOR";
            sheet.Cells[i, 5].Value = "VENDEDOR";
            sheet.Cells[i, 6].Value = "GESTOR";
            sheet.Cells[i, 7].Value = "LINHA";

            i++;

            LinhaDevolucao.ForEach(Pedido =>
            {
                sheet.Cells[i, 1].Value = Pedido.Empresa;
                sheet.Cells[i, 2].Value = Pedido.D1Nfori;
                sheet.Cells[i, 3].Value = Pedido.Nome;
                sheet.Cells[i, 4].Value = Pedido.Total;
                sheet.Cells[i, 5].Value = Pedido.Login;
                sheet.Cells[i, 6].Value = Pedido.Gestor;
                sheet.Cells[i, 7].Value = Pedido.Linha;

                i++;
            });


            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

            //sheet.Protection.IsProtected = true;
            using MemoryStream stream = new MemoryStream();
            package.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"RelatorioComissoesRH{Mes}{Ano}{usuario.Integrante.ToUpper().Replace(".","")}.xlsx");

        }
    }
}
