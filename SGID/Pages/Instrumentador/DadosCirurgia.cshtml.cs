using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Cirurgias;
using SGID.Models.Denuo;
using SGID.Models.DTO;
using SGID.Models.Inter;

namespace SGID.Pages.Instrumentador
{
    [Authorize]
    public class DadosCirurgiaModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext INTER { get; set; }
        private TOTVSDENUOContext DENUO { get; set; }
        public Agendamentos Agendamento { get; set; }

        public List<Produto> Produtos { get; set; }
        public List<Patrimonio> Patrimonios { get; set; }
        public List<Produto> Avulsos { get; set; }

        private readonly IWebHostEnvironment _WEB;

        public List<string> SearchProduto { get; set; } = new List<string>();

        public List<string> Procedimento { get; set; } = new List<string>();

        public DadosCirurgiaModel(ApplicationDbContext sgid, IWebHostEnvironment wEB,TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            SGID = sgid;
            _WEB = wEB;
            INTER = inter;
            DENUO = denuo;
        }

        public void OnGet(int id)
        {
           Agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == id);
            var codigos = SGID.ProdutosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id)
                    .Select(x => new
                    {
                        produto = x.CodigoProduto,
                        unidade = x.Quantidade,
                        codTabela = x.CodigoTabela,
                        ValorUnidade = x.ValorUnitario,
                        valor = x.ValorTotal
                    }).ToList();

            var Patrimonio = SGID.PatrimoniosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id)
                .Select(x => new { x.Patrimonio, x.Codigo }).ToList();

            var avulsos = SGID.AvulsosAgendamento.Where(x => x.AgendamentoId == Agendamento.Id)
                .Select(x => new
                {
                    Produto = x.Produto,
                    Quant = x.Quantidade
                }).ToList();

            Produtos = new List<Produto>();
            Patrimonios = new List<Patrimonio>();
            Avulsos = new List<Produto>();
            if (Agendamento.Empresa == "01")
            {
                //J&J
                codigos.ForEach(x =>
                {
                    var produto = (from SB10 in INTER.Sb1010s
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

                    var preco = (from DA10 in INTER.Da1010s
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

                Patrimonio.ForEach(x =>
                {
                    var view = (from PA10 in INTER.Pa1010s
                                join PAC in INTER.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                from c in sr.DefaultIfEmpty()
                                join SA10 in INTER.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
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
                    var produto = (from SB10 in INTER.Sb1010s
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
                    };

                    Avulsos.Add(ViewProduto);
                });

                SearchProduto = INTER.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1" && x.B1Tipo != "KT" && x.B1Comerci == "C").Select(x => x.B1Desc).Distinct().ToList();
                Procedimento = SGID.Procedimentos.Where(x => x.Bloqueado == 0 && x.Empresa == "01").Select(x => x.Nome).ToList();
            }
            else
            {
                  //FLOWMED
                codigos.ForEach(x =>
                {
                    var produto = (from SB10 in DENUO.Sb1010s
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

                Patrimonio.ForEach(x =>
                {
                    var view = (from PA10 in DENUO.Pa1010s
                                join PAC in DENUO.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                from c in sr.DefaultIfEmpty()
                                join SA10 in DENUO.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
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
                    var produto = (from SB10 in DENUO.Sb1010s
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

                    var preco = (from DA10 in DENUO.Da1010s
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

                SearchProduto = DENUO.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1" && x.B1Tipo != "KT" && x.B1Comerci == "C").Select(x => x.B1Desc).Distinct().ToList();
                Procedimento = SGID.Procedimentos.Where(x => x.Bloqueado == 0 && x.Empresa == "03").Select(x => x.Nome).ToList();
            }
        }

        public IActionResult OnPostAsync(int Id,string Codigo,string NomePaciente,string NomeMedico,string NomeCliente,int Status,
            DateTime DataCirurgia,string Hora,string Procedimento,string Obs, IFormCollection Anexos01, IFormCollection Anexos02, 
            IFormCollection Anexos03, IFormCollection Anexos04, IFormCollection Anexos05, List<Produto> Produtos,string Hospital,
            string Especialidade,string Localidade,string Semana)
        {

            var dados = new DadosCirurgia
            {
                DataCriacao = DateTime.Now,
                Codigo = Codigo,
                NomePaciente = NomePaciente,
                NomeMedico = NomeMedico,
                NomeCliente = NomeCliente,
                Hospital = Hospital,
                Status = Status,
                DataCirurgia = DataCirurgia,
                ProcedimentosExec = Procedimento,
                ObsIntercorrencia = Obs,
                AgendamentoId = Id,
                Especialidade = Especialidade,
                Localidade = Localidade,
                Semana = Semana
            };

            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == Id);

            if(Status == 2)
            {
                agendamento.StatusInstrumentador = 4;
            }
            else
            {
                agendamento.StatusInstrumentador = 3;
            }

            

            SGID.DadosCirurgias.Add(dados);
            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

            string Pasta = $"{_WEB.WebRootPath}/AnexosDadosCirurgia";

            if (!Directory.Exists(Pasta))
            {
                Directory.CreateDirectory(Pasta);
            }

            //Anexos
            #region Anexos

            foreach (var anexo in Anexos01.Files)
            {
                var anexoAgenda = new AnexosDadosCirurgia
                {
                    DadosCirurgiaId = dados.Id,
                    AnexoCam = $"{dados.Id}01.{anexo.FileName.Split(".").Last()}"
                };

                string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";
                using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                {
                    anexo.CopyTo(fileStream);
                }

                SGID.AnexosDadosCirurgias.Add(anexoAgenda);
                SGID.SaveChanges();
            }
            foreach (var anexo in Anexos02.Files)
            {
                var anexoAgenda = new AnexosDadosCirurgia
                {
                    DadosCirurgiaId = dados.Id,
                    AnexoCam = $"{dados.Id}02.{anexo.FileName.Split(".").Last()}"
                };

                string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";
                using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                {
                    anexo.CopyTo(fileStream);
                }

                SGID.AnexosDadosCirurgias.Add(anexoAgenda);
                SGID.SaveChanges();
            }
            foreach (var anexo in Anexos03.Files)
            {
                var anexoAgenda = new AnexosDadosCirurgia
                {
                    DadosCirurgiaId = dados.Id,
                    AnexoCam = $"{dados.Id}03.{anexo.FileName.Split(".").Last()}"
                };

                string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";
                using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                {
                    anexo.CopyTo(fileStream);
                }

                SGID.AnexosDadosCirurgias.Add(anexoAgenda);
                SGID.SaveChanges();
            }
            foreach (var anexo in Anexos04.Files)
            {
                var anexoAgenda = new AnexosDadosCirurgia
                {
                    DadosCirurgiaId = dados.Id,
                    AnexoCam = $"{dados.Id}04.{anexo.FileName.Split(".").Last()}"
                };

                string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";
                using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                {
                    anexo.CopyTo(fileStream);
                }

                SGID.AnexosDadosCirurgias.Add(anexoAgenda);
                SGID.SaveChanges();
            }
            foreach (var anexo in Anexos05.Files)
            {
                var anexoAgenda = new AnexosDadosCirurgia
                {
                    DadosCirurgiaId = dados.Id,
                    AnexoCam = $"{dados.Id}05.{anexo.FileName.Split(".").Last()}"
                };

                string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";
                using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                {
                    anexo.CopyTo(fileStream);
                }

                SGID.AnexosDadosCirurgias.Add(anexoAgenda);
                SGID.SaveChanges();
            }

            #endregion

            Produtos.ForEach(produto =>
            {
                var ProdXAgenda = new DadosCirugiasProdutos
                {
                    DadosCirurgiaId = dados.Id, 
                    Quantidade = produto.Und,
                    Produto = produto.Item
                };

                SGID.DadosCirugiasProdutos.Add(ProdXAgenda);
                SGID.SaveChanges();

            });

            return LocalRedirect("/dashboard/0");
        }
    }
}
