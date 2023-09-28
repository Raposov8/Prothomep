using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Cirurgias;
using SGID.Models.Denuo;
using SGID.Models.DTO;

namespace SGID.Pages.Logistica
{
    public class CheckAgendamentoModel : PageModel
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

        public CheckAgendamentoModel(ApplicationDbContext sgid,TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            SGID = sgid;
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
                        valor = x.ValorTotal
                    }).ToList();

            var Patrimonio = SGID.PatrimoniosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id)
                .Select(x => new { x.Patrimonio,x.Codigo }).ToList();

            var avulsos = SGID.AvulsosAgendamento.Where(x => x.AgendamentoId == Agendamento.Id)
                .Select(x => new
                {
                    Produto = x.Produto,
                    Quant = x.Quantidade
                }).ToList();

            var anexos = SGID.AnexosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id).ToList();

            if (anexos.Count > 0)
            {
                anexos.ForEach(x => Anexos.Add(x.AnexoCam));
            }

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
                Crm = ProtheusInter.Sa1010s.FirstOrDefault(x => x.A1Nreduz == Agendamento.Medico).A1Crm;
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
                        PrcUnid = preco,
                        SegUnd = produto.B1Um,
                        VlrTotal = x.valor,
                    };

                    Produtos.Add(ViewProduto);
                });

                Patrimonio.ForEach(x =>
                {
                    var view = (from PA10 in ProtheusInter.Pa1010s
                                join PAC in ProtheusInter.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                from c in sr.DefaultIfEmpty()
                                join SA10 in ProtheusInter.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                                from a in st.DefaultIfEmpty()
                                where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                                && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == x.Patrimonio
                                select new Patrimonio
                                {
                                    Descri = PA10.Pa1Despat,
                                    KitBas = PA10.Pa1Kitbas,
                                    Codigo = x.Codigo
                                }).First();


                    Patrimonios.Add(view);
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

                    var preco = (from DA10 in ProtheusInter.Da1010s
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
                    };

                    Avulsos.Add(ViewProduto);
                });
            }
            else
            {
                //Denuo
                Crm = ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.A1Nreduz == Agendamento.Medico).A1Crm;
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

                    var preco = (from DA10 in ProtheusDenuo.Da1010s
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
                        PrcUnid = preco,
                        SegUnd = produto.B1Um,
                        VlrTotal = x.valor,
                    };

                    Produtos.Add(ViewProduto);
                });

                Patrimonio.ForEach(x =>
                {
                    var view = (from PA10 in ProtheusDenuo.Pa1010s
                                join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                from c in sr.DefaultIfEmpty()
                                join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                                from a in st.DefaultIfEmpty()
                                where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                                && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == x.Patrimonio
                                && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                                select new Patrimonio
                                {
                                    Descri = PA10.Pa1Despat,
                                    KitBas = PA10.Pa1Kitbas,
                                    Codigo = x.Codigo
                                }).First();


                    Patrimonios.Add(view);
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
                    };

                    Avulsos.Add(ViewProduto);
                });
            }
        }

        public IActionResult OnPost(int id,List<Produto> Avulsos, List<Produto> Produtos, List<Patrimonio> Patris)
        {

            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == id);


            Produtos.ForEach(Prod =>
            {

                var checkAgen = new AgendamentoCheck { AgendamentoId = agendamento.Id, Codigo = Prod.Item, Descricao = Prod.Produtos, Quantidade = Prod.Und, Entregue = Prod.Check =="true" ? true:false };

                SGID.AgendamentoChecks.Add(checkAgen);
                SGID.SaveChanges();

            });

            Patris.ForEach(Patr =>
            {
                var checkAgen = new AgendamentoCheck { AgendamentoId = agendamento.Id, Codigo = Patr.KitBas, Descricao = Patr.Descri, Quantidade = 1, Entregue = Patr.Check == "true" ? true : false };

                SGID.AgendamentoChecks.Add(checkAgen);
                SGID.SaveChanges();
            });

            Avulsos.ForEach(Avul =>
            {
                var checkAgen = new AgendamentoCheck { AgendamentoId = agendamento.Id, Codigo = Avul.Item, Descricao = Avul.Produtos, Quantidade = Avul.Und, Entregue = Avul.Check == "true" ? true : false };

                SGID.AgendamentoChecks.Add(checkAgen);
                SGID.SaveChanges();
            });


            agendamento.StatusLogistica = 3;
            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();


            return LocalRedirect("/logistica/listarlogistica/3");
        }
    }
}
