using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Cirurgias;
using SGID.Models.DTO;
using SGID.Models.Denuo;
using OPMEnexo;
using SGID.Models.Inter;

namespace SGID.Pages.Cotacoes
{
    [Authorize(Roles = "Admin,GestorComercial,Comercial,Diretoria")]
    public class AceitarCotacoesModel : PageModel
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
        public AceitarCotacoesModel(ApplicationDbContext sgid, IWebHostEnvironment wEB, TOTVSDENUOContext denuo, TOTVSINTERContext inter)
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
                        valor = x.ValorTotal
                    }).ToList();

            var Patrimonio = SGID.PatrimoniosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id)
                .Select(x => x.Patrimonio).ToList();

            var avulsos = SGID.AvulsosAgendamento.Where(x => x.AgendamentoId == Agendamento.Id)
                .Select(x => new
                {
                    Produto = x.Produto,
                    Quant = x.Quantidade
                }).ToList();

            var anexos = SGID.AnexosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id).ToList();

            if(anexos.Count > 0)
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
                                && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == x
                                select new Patrimonio
                                {
                                    Descri = PA10.Pa1Despat,
                                    KitBas = PA10.Pa1Kitbas,
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

                SearchPatri = ProtheusInter.Pa1010s.Where(x => x.DELET != "*" && x.Pa1Msblql != "1").Select(x => x.Pa1Despat).Distinct().ToList();

                SearchProduto = ProtheusInter.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1").Select(x => x.B1Desc).Distinct().ToList();
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
                                && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == x
                                && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                                select new Patrimonio
                                {
                                    Descri = PA10.Pa1Despat,
                                    KitBas = PA10.Pa1Kitbas,
                                }).First();


                    Patrimonios.Add(view);
                });

                SearchPatri = (from PA10 in ProtheusDenuo.Pa1010s
                               join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                               from c in sr.DefaultIfEmpty()
                               join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                               from a in st.DefaultIfEmpty()
                               where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                               && c.DELET != "*" && a.DELET != "*"
                               && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                               select PA10.Pa1Despat
                                   ).Distinct().ToList();

                SearchProduto = ProtheusDenuo.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1").Select(x => x.B1Desc).Distinct().ToList();
                
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

        public IActionResult OnPostAsync(int id,string NomePaciente,DateTime? DataCirurgia,int Urgencia
            ,string NomeMedico,string CRM,string NomeHospital,string NomeVendedor,string Obs,
            List<Produto> Avulsos, List<Produto> Produtos,List<Patrimonio> Patris ,string CodTabela,DateTime? Entrega)
        {
            try
            {
                var AgendamentoProduto = SGID.ProdutosAgendamentos.Where(x => x.AgendamentoId == id).ToList();

                if (!string.IsNullOrEmpty(Obs) && !string.IsNullOrWhiteSpace(Obs))
                {
                    var Observacao = new ObsAgendamento { AgendamentoId = id, User = User.Identity.Name.Split("@")[0].ToUpper(), Obs = Obs, DataCriacao = DateTime.Now };

                    SGID.ObsAgendamentos.Add(Observacao);
                    SGID.SaveChanges();
                }

                AgendamentoProduto.ForEach(produto => 
                {
                    var produtoUpdate = Produtos.FirstOrDefault(x => x.Item == produto.CodigoProduto);

                    if (produtoUpdate != null)
                    {

                        produto.Quantidade = produtoUpdate.Und;
                        produto.ValorTotal = produtoUpdate.Und * produtoUpdate.PrcUnid;

                        SGID.ProdutosAgendamentos.Update(produto);
                        SGID.SaveChanges();

                        Produtos.Remove(produtoUpdate);
                    }
                    else
                    {
                        SGID.ProdutosAgendamentos.Remove(produto);
                        SGID.SaveChanges();
                    }
                });

                Produtos.ForEach(produto => 
                {
                    var ProdXAgenda = new ProdutosAgendamentos
                    {
                        AgendamentoId = id,
                        CodigoProduto = produto.Item,
                        Quantidade = produto.Und,
                        ValorTotal = produto.Und * produto.PrcUnid,
                        CodigoTabela = CodTabela
                    };

                    SGID.ProdutosAgendamentos.Add(ProdXAgenda);
                    SGID.SaveChanges();
                });

                var AgendamentoPatris = SGID.PatrimoniosAgendamentos.Where(x => x.AgendamentoId == id).ToList();

                AgendamentoPatris.ForEach(produto => 
                {
                    var patriUpdate = Patris.FirstOrDefault(x => x.Descri == produto.Patrimonio);

                    SGID.PatrimoniosAgendamentos.Remove(produto);
                    SGID.SaveChanges();

                });

                Patris.ForEach(patri => 
                {
                    var ProdXAgenda = new PatrimonioAgendamento
                    {
                        AgendamentoId = id,
                        Patrimonio = patri.Descri
                    };

                    SGID.PatrimoniosAgendamentos.Add(ProdXAgenda);
                    SGID.SaveChanges();
                });

                var AgendamentoAvulsos = SGID.AvulsosAgendamento.Where(x => x.AgendamentoId == id).ToList();

                AgendamentoAvulsos.ForEach(avus => 
                {
                    var Avulso = Avulsos.FirstOrDefault(c => c.Item == avus.Produto);

                    if (Avulso != null)
                    {
                        avus.Quantidade = Avulso.Und;

                        SGID.AvulsosAgendamento.Update(avus);
                        SGID.SaveChanges();

                    }
                    else
                    {
                        SGID.AvulsosAgendamento.Remove(avus);
                        SGID.SaveChanges();
                    }
                });

                Avulsos.ForEach(avulso => 
                {
                    var agendamento = new AvulsosAgendamento
                    {
                        AgendamentoId = id,
                        Produto = avulso.Item,
                        Quantidade = avulso.Und
                    };


                    SGID.AvulsosAgendamento.Add(agendamento);
                    SGID.SaveChanges();

                });

                var Agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == id);

                Agendamento.DataCirurgia = DataCirurgia;
                Agendamento.DataAlteracao = DateTime.Now;
                Agendamento.DataEntrega = Entrega;

                if (Agendamento.StatusPedido == 2)
                {
                    Agendamento.StatusCotacao = 1;
                    Agendamento.StatusPedido = 5;
                    Agendamento.UsuarioComercial = User.Identity.Name.Split("@")[0].ToUpper();
                }



                SGID.Agendamentos.Update(Agendamento);
                SGID.SaveChanges();

                return LocalRedirect("/cotacoes/DashBoardCotacoes/0");
            }catch(Exception E)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(E, SGID, "AceitarCotacoes", user);
                return LocalRedirect("/cotacoes/DashBoardCotacoes/0");
            }
        }
    }
}
