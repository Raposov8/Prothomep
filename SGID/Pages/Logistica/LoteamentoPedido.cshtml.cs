using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;
using SGID.Data;
using SGID.Models.Denuo;
using SGID.Models.Inter;
using SGID.Data.Models;
using SGID.Models.Cirurgias;
using SGID.Models.DTO;

namespace SGID.Pages.Logistica
{
    public class LoteamentoPedidoModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        public Agendamentos Agendamento { get; set; }
        public string Crm { get; set; }
        public List<string> Anexos { get; set; } = new List<string>();
        public List<Produto> Produtos { get; set; }

        public List<Produto> Avulsos { get; set; }
        public List<Patrimonio> Patrimonios { get; set; }
        public List<Procedimento> Procedimentos { get; set; }
        public List<ObsAgendamento> Observacoes { get; set; } = new List<ObsAgendamento>();

        public List<string> SearchPatri { get; set; } = new List<string>();
        public List<string> SearchProduto { get; set; } = new List<string>();

        private readonly IWebHostEnvironment _WEB;
        public LoteamentoPedidoModel(ApplicationDbContext sgid, IWebHostEnvironment wEB, TOTVSDENUOContext denuo, TOTVSINTERContext inter)
        {
            SGID = sgid;
            _WEB = wEB;
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
        }

        public void OnGet(int id)
        {
            Agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == id);
            Procedimentos = SGID.Procedimentos.Where(x => x.Bloqueado == 0).OrderBy(x => x.Nome).ToList();
            var data = Convert.ToInt32(DateTime.Now.ToString("yyyy/MM/dd").Replace("/", ""));

            var codigos = SGID.ProdutosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id)
                    .Select(x => new
                    {
                        produto = x.CodigoProduto,
                        unidade = x.Quantidade,
                        codTabela = x.CodigoTabela,
                        ValorUnidade = x.ValorUnitario,
                        valor = x.ValorTotal
                    }).ToList();

            var avulsos = SGID.AvulsosAgendamento.Where(x => x.AgendamentoId == Agendamento.Id)
                .Select(x => new
                {
                    Produto = x.Produto,
                    Quant = x.Quantidade,
                    Lote = x.Lote
                }).ToList();

            Observacoes = SGID.ObsAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id).ToList();

            if (Observacoes.Count == 0)
            {
                Observacoes.Add(new ObsAgendamento { User = "Não ha Registro", Obs = "Não ha Registro" });
            }


            Produtos = new List<Produto>();
            Patrimonios = new List<Patrimonio>();
            Avulsos = new List<Produto>();
            if (Agendamento.Empresa == "01")
            {
                //Intermedic
                Crm = ProtheusInter.Sa1010s.FirstOrDefault(x => x.A1Nome == Agendamento.Medico).A1Crm;
                codigos.ForEach(x =>
                {
                    var produto = (from SB10 in ProtheusInter.Sb1010s
                                   where SB10.B1Cod == x.produto.ToUpper() && SB10.B1Msblql != "1"
                                   && SB10.DELET != "*"
                                   select new
                                   {
                                       SB10.B1Cod,
                                       SB10.B1Msblql,
                                       SB10.B1Solicit,
                                       SB10.B1Desc,
                                       SB10.B1Fabric,
                                       SB10.B1Tipo,
                                       SB10.B1Lotesbp,
                                       SB10.B1Um,
                                       SB10.B1Reganvi,
                                       SB10.B1Xtuss
                                   }).FirstOrDefault();

                    var preco = (from DA10 in ProtheusInter.Da1010s
                                 where DA10.DELET != "*" && DA10.Da1Codtab == x.codTabela && DA10.Da1Codpro == x.produto.ToUpper()
                                 select DA10.Da1Prcven).FirstOrDefault();


                    var ViewProduto = new Produto
                    {
                        Item = produto.B1Cod,
                        Licit = produto.B1Solicit,
                        Produtos = produto.B1Desc,
                        Tuss = produto.B1Xtuss,
                        Anvisa = produto.B1Reganvi,
                        Marca = produto.B1Fabric,
                        Und = x.unidade,
                        PrcUnid = x.ValorUnidade,
                        SegUnd = produto.B1Um,
                        VlrTotal = x.valor,
                    };

                    Produtos.Add(ViewProduto);
                });

                avulsos.ForEach(x =>
                {
                    var produto = (from SB10 in ProtheusInter.Sb1010s
                                   where SB10.B1Cod == x.Produto.ToUpper() && SB10.B1Msblql != "1"
                                   && SB10.DELET != "*"
                                   select new
                                   {
                                       SB10.B1Cod,
                                       SB10.B1Msblql,
                                       SB10.B1Solicit,
                                       SB10.B1Desc,
                                       SB10.B1Fabric,
                                       SB10.B1Tipo,
                                       SB10.B1Lotesbp,
                                       SB10.B1Um,
                                       SB10.B1Reganvi,
                                       SB10.B1Xtuss
                                   }).FirstOrDefault();


                    var ViewProduto = new Produto
                    {
                        Item = produto.B1Cod,
                        Licit = produto.B1Solicit,
                        Produtos = produto.B1Desc,
                        Tuss = produto.B1Xtuss,
                        Anvisa = produto.B1Reganvi,
                        Marca = produto.B1Fabric,
                        Und = x.Quant,
                        SegUnd = produto.B1Um,
                        Lote = x.Lote
                    };

                    var local = new[] { "01", "30", "80" };

                    ViewProduto.Lotes = ProtheusInter.Sbf010s.Where(x => x.DELET != "*" && local.Contains(x.BfLocal) && x.BfProduto == produto.B1Cod).Select(x => x.BfLotectl).ToList();

                    Avulsos.Add(ViewProduto);
                });

            }
            else
            {
                //Denuo
                Crm = ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.A1Nome == Agendamento.Medico).A1Crm;
                codigos.ForEach(x =>
                {
                    var produto = (from SB10 in ProtheusDenuo.Sb1010s
                                   where SB10.B1Cod == x.produto.ToUpper() && SB10.B1Msblql != "1"
                                   && SB10.DELET != "*"
                                   select new
                                   {
                                       SB10.B1Cod,
                                       SB10.B1Msblql,
                                       SB10.B1Solicit,
                                       SB10.B1Desc,
                                       SB10.B1Fabric,
                                       SB10.B1Tipo,
                                       SB10.B1Lotesbp,
                                       SB10.B1Um,
                                       SB10.B1Reganvi,
                                       SB10.B1Xtuss
                                   }).FirstOrDefault();


                    var ViewProduto = new Produto
                    {
                        Item = produto.B1Cod,
                        Licit = produto.B1Solicit,
                        Produtos = produto.B1Desc,
                        Tuss = produto.B1Xtuss,
                        Anvisa = produto.B1Reganvi,
                        Marca = produto.B1Fabric,
                        Und = x.unidade,
                        PrcUnid = x.ValorUnidade,
                        SegUnd = produto.B1Um,
                        VlrTotal = x.valor,
                    };

                    Produtos.Add(ViewProduto);
                });

                avulsos.ForEach(x =>
                {
                    var produto = (from SB10 in ProtheusDenuo.Sb1010s
                                   where SB10.B1Cod == x.Produto.ToUpper() && SB10.B1Msblql != "1"
                                   && SB10.DELET != "*"
                                   select new
                                   {
                                       SB10.B1Cod,
                                       SB10.B1Msblql,
                                       SB10.B1Solicit,
                                       SB10.B1Desc,
                                       SB10.B1Fabric,
                                       SB10.B1Tipo,
                                       SB10.B1Lotesbp,
                                       SB10.B1Um,
                                       SB10.B1Reganvi,
                                       SB10.B1Xtuss
                                   }).FirstOrDefault();

                    var preco = (from DA10 in ProtheusDenuo.Da1010s
                                 where DA10.DELET != "*" && DA10.Da1Codtab == Agendamento.CodTabela && DA10.Da1Codpro == x.Produto.ToUpper()
                                 select DA10.Da1Prcven).FirstOrDefault();


                    var ViewProduto = new Produto
                    {
                        Item = produto.B1Cod,
                        Licit = produto.B1Solicit,
                        Produtos = produto.B1Desc,
                        Tuss = produto.B1Xtuss,
                        Anvisa = produto.B1Reganvi,
                        Marca = produto.B1Fabric,
                        Und = x.Quant,
                        PrcUnid = preco,
                        SegUnd = produto.B1Um,
                        VlrTotal = preco,
                        Lote = x.Lote
                    };

                    var local = new[] { "01", "30",  "80" };

                    ViewProduto.Lotes = ProtheusDenuo.Sbf010s.Where(x => x.DELET != "*" && local.Contains(x.BfLocal) && x.BfProduto == produto.B1Cod).Select(x=> x.BfLotectl).ToList();

                    Avulsos.Add(ViewProduto);
                });
            }
        }

        public IActionResult OnPostAsync(int id, string Empresa, string NomePaciente, DateTime? DataCirurgia, int Urgencia
            , string NomeMedico, string CRM, string NomeHospital, string NomeVendedor, string Obs,
            List<Produto> Avulsos, List<Produto> Produtos, List<Patrimonio> Patris, string CodTabela, DateTime? Entrega, int Reprovado)
        {
            try
            {
                if (!string.IsNullOrEmpty(Obs) && !string.IsNullOrWhiteSpace(Obs))
                {
                    var Observacao = new ObsAgendamento { AgendamentoId = id, User = User.Identity.Name.Split("@")[0].ToUpper(), Obs = Obs, DataCriacao = DateTime.Now };

                    SGID.ObsAgendamentos.Add(Observacao);
                    SGID.SaveChanges();
                }

                #region Avulsos

                var AgendamentoAvulsos = SGID.AvulsosAgendamento.Where(x => x.AgendamentoId == id).ToList();

                AgendamentoAvulsos.ForEach(avus =>
                {
                    var Avulso = Avulsos.FirstOrDefault(c => c.Item == avus.Produto);

                    avus.Quantidade = Avulso.Und;
                    avus.Lote = Avulso.Lote;

                    SGID.AvulsosAgendamento.Update(avus);
                    SGID.SaveChanges();
                });

                #endregion

                var Agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == id);

                Agendamento.UsuarioLoteamento = User.Identity.Name.Split("@")[0].ToUpper();
                Agendamento.DataLoteamento = DateTime.Now;
                //Agendamento.StatusLogistica = 7;

                SGID.Agendamentos.Update(Agendamento);
                SGID.SaveChanges();

                return LocalRedirect($"/logistica/listarlogistica/{Empresa}/2");

            }
            catch (Exception E)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(E, SGID, "LoteamentoPedido", user);
                return LocalRedirect($"/logistica/listarlogistica/{Empresa}/2");
            }
        }
    }
}
