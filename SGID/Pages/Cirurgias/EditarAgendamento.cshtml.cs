using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;
using SGID.Data;
using SGID.Models.Cirurgias;
using SGID.Models.DTO;
using SGID.Data.Models;
using SGID.Models.Denuo;
using Microsoft.AspNetCore.Authorization;
using System.Data;
using SGID.Models.Inter;

namespace SGID.Pages.Cirurgias
{
    [Authorize]
    public class EditarAgendamentoModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private readonly IWebHostEnvironment _WEB;

        public Agendamentos Agendamento { get; set; }
        public List<Produto> Produtos { get; set; }
        public List<Patrimonio> Patrimonios { get; set; }
        public List<ObsAgendamento> Observacoes { get; set; } = new List<ObsAgendamento>();
        public NovoAgendamento Novo { get; set; } = new NovoAgendamento();
        public string Anexo1 { get; set; }
        public string Anexo2 { get; set; }
        public string Anexo3 { get; set; }
        public string Anexo4 { get; set; }
        public string Anexo5 { get; set; }

        public List<string> SearchProduto { get; set; } = new List<string>();

        public EditarAgendamentoModel(ApplicationDbContext sgid, TOTVSINTERContext protheus, TOTVSDENUOContext denuo, IWebHostEnvironment wEB)
        {
            SGID = sgid;
            ProtheusInter = protheus;
            ProtheusDenuo = denuo;
            _WEB = wEB;
        }
        public void OnGet(int id)
        {
            Agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == id);
            Novo = new NovoAgendamento();
            var data = Convert.ToInt32(DateTime.Now.ToString("yyyy/MM/dd").Replace("/", ""));

            var codigos = SGID.ProdutosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id)
                    .Select(x => new
                    {
                        produto = x.CodigoProduto,
                        unidade = x.Quantidade,
                        valorUnitario = x.ValorUnitario,
                        valor = x.ValorTotal,
                        Tabela = x.CodigoTabela
                    }).ToList();

            var Patrimonio = SGID.PatrimoniosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id)
                .Select(x => x.Patrimonio).ToList();

            var anexos = SGID.AnexosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id).ToList();

            anexos.ForEach(anexo =>
            {
                switch (anexo.NumeroAnexo)
                {
                    case "01": Anexo1 = "S"; break;
                    case "02": Anexo2 = "S"; break;
                    case "03": Anexo3 = "S"; break;
                    case "04": Anexo4 = "S"; break;
                    case "05": Anexo5 = "S"; break;
                }
            });

            Observacoes = SGID.ObsAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id).ToList();

            if(Observacoes.Count == 0)
            {
                Observacoes.Add(new ObsAgendamento { User = "Não ha Registro", Obs = "Não ha Registro" });
            }

            Produtos = new List<Produto>();
            Patrimonios = new List<Patrimonio>();

            if (Agendamento.Empresa == "01")
            {
                //Intermedic
                Novo = new NovoAgendamento
                {
                    Clientes = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Msblql != "1" && (x.A1Clinter == "C" || x.A1Clinter == "H" || x.A1Clinter == "M")).OrderBy(x => x.A1Nome).Select(x => x.A1Nreduz).ToList(),
                    Convenio = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "C" && x.A1Msblql != "1").OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                    Medico = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Crm != "").OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                    Intrumentador = ProtheusInter.Pah010s.Where(x => x.DELET != "*" && x.PahMsblql != "1").OrderBy(x => x.PahNome).Select(x => x.PahNome).ToList(),
                    Hospital = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "H" && x.A1Msblql != "1").OrderBy(x => x.A1Nome).Select(x => x.A1Nreduz).ToList(),
                    Procedimentos = SGID.Procedimentos.Where(x => x.Bloqueado == 0 && x.Empresa == "01").ToList(),
                    Patrimonios = ProtheusInter.Pa1010s.Where(x => x.DELET != "*" && x.Pa1Msblql != "1").Select(x => x.Pa1Despat).Distinct().ToList(),
                    Tabelas = ProtheusInter.Da0010s.Where(x => x.DELET != "*").Select(x => x.Da0Descri).ToList(),
                    Condicoes = ProtheusInter.Se4010s.Where(x => x.DELET != "*" && x.E4Msblql != "1").Select(x => x.E4Descri).ToList(),
                    Vendedores = ProtheusInter.Sa3010s.Where(x => x.DELET != "*" && x.A3Msblql != "1").Select(x => x.A3Nreduz).ToList(),
                };

                SearchProduto = ProtheusInter.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1" && x.B1Tipo != "KT" && x.B1Comerci=="C").Select(x => x.B1Cod + " " + x.B1Desc).Distinct().ToList();

                codigos.ForEach(x =>
                {
                    //Intermedic
                    var produto = (from SB10 in ProtheusInter.Sb1010s
                                   where SB10.B1Cod == x.produto && SB10.B1Msblql != "1"
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
                        PrcUnid = x.valorUnitario,
                        SegUnd = produto.B1Um,
                        VlrTotal = x.valor,
                        TipoOp = produto.B1Tipo,
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
            }
            else
            {
                //Denuo

                Novo = new NovoAgendamento
                {
                    Clientes = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Msblql != "1" && (x.A1Clinter == "C" || x.A1Clinter == "H" || x.A1Clinter == "M")).OrderBy(x => x.A1Nome).Select(x => x.A1Nreduz).ToList(),
                    Convenio = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "C" && x.A1Msblql != "1").OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                    Medico = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Crm != "").OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                    Intrumentador = ProtheusDenuo.Pah010s.Where(x => x.DELET != "*" && x.PahMsblql != "1").OrderBy(x => x.PahNome).Select(x => x.PahNome).ToList(),
                    Hospital = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "H" && x.A1Msblql != "1").OrderBy(x => x.A1Nome).Select(x => x.A1Nreduz).ToList(),
                    Procedimentos = SGID.Procedimentos.Where(x => x.Bloqueado == 0 && x.Empresa == "03").ToList(),
                    Patrimonios = (from PA10 in ProtheusDenuo.Pa1010s
                                   join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                   from c in sr.DefaultIfEmpty()
                                   join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                                   from a in st.DefaultIfEmpty()
                                   where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                                   && c.DELET != "*" && a.DELET != "*"
                                   && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                                   select PA10.Pa1Despat
                                   ).Distinct().ToList(),
                    Tabelas = ProtheusDenuo.Da0010s.Where(x => x.DELET != "*").Select(x => x.Da0Descri).ToList(),
                    Condicoes = ProtheusDenuo.Se4010s.Where(x => x.DELET != "*" && x.E4Msblql != "1").Select(x => x.E4Descri).ToList(),
                    Vendedores = ProtheusDenuo.Sa3010s.Where(x => x.DELET != "*" && x.A3Msblql != "1").Select(x => x.A3Nreduz).ToList(),
                };

                SearchProduto = ProtheusDenuo.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1" && x.B1Tipo != "KT" && x.B1Comerci=="C").Select(x => x.B1Cod + " " + x.B1Desc).Distinct().ToList();

                codigos.ForEach(x =>
                {
                    //Denuo
                    var produto = (from SB10 in ProtheusDenuo.Sb1010s
                                   where SB10.B1Cod == x.produto && SB10.B1Msblql != "1"
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
                            PrcUnid = x.valorUnitario,
                            SegUnd = produto.B1Um,
                            VlrTotal = x.valor,
                            TipoOp = produto.B1Tipo,
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
                                select new Patrimonio
                                {
                                    Descri = PA10.Pa1Despat,
                                    KitBas = PA10.Pa1Kitbas,
                                }).First();


                    Patrimonios.Add(view);
                });
            }
        }

        public IActionResult OnPostAsync(int Id,string Empresa, string Cliente, string CondPag, string Tabela, string Vendedor,
            string Medico, string Matricula, string Paciente, string Convenio, string Instrumentador, string Hospital,
            DateTime? DataAgendamento, DateTime DataAutorizacao, string NumAutorizacao, string Senha, int Tipo, int Autorizado, string Obs,
            List<Produto> Produtos, List<Patrimonio> Patris, IFormFile Anexos01, IFormFile Anexos02, IFormFile Anexos03, IFormFile Anexos04, IFormFile Anexos05,
            string CodCondPag,string CodHospital,string CodConvenio,string CodTabela,DateTime? DataEntrega, DateTime? DataRetirada, string Procedimento, string Indicacao)
        {
            try
            {
                var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == Id);

                agendamento.DataAlteracao = DateTime.Now;
                agendamento.CodHospital = CodHospital;
                agendamento.Hospital = Hospital;
                agendamento.Cliente = Cliente;
                agendamento.CodCondPag = CodCondPag;
                agendamento.CondPag = CondPag;
                agendamento.CodTabela = CodTabela;
                agendamento.Tabela = Tabela;
                agendamento.Vendedor = Vendedor;
                agendamento.Medico = Medico;
                agendamento.Matricula = Matricula;
                agendamento.Paciente = Paciente;
                agendamento.CodConvenio = CodConvenio;
                agendamento.Convenio = Convenio;
                agendamento.Instrumentador = Instrumentador;
                agendamento.DataCirurgia = DataAgendamento;
                agendamento.DataAutorizacao = DataAutorizacao;
                agendamento.NumAutorizacao = NumAutorizacao;
                agendamento.Senha = Senha;
                agendamento.Tipo = Tipo;
                agendamento.Procedimento = Procedimento;
                agendamento.Indicacao = Indicacao;
                agendamento.Autorizado = Autorizado;
                agendamento.Observacao = Obs;

                agendamento.DataEntrega = DataEntrega;
                agendamento.DataRetorno = DataRetirada;

                agendamento.UsuarioAlterar = User.Identity.Name.Split("@")[0].ToUpper();

                if (Empresa == "01")
                {
                    //Intermedic
                    agendamento.VendedorLogin = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Nreduz == Vendedor).A3Xlogin;
                }
                else
                {
                    //Denuo
                    agendamento.VendedorLogin = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Nreduz == Vendedor).A3Xlogin;
                }

                SGID.Agendamentos.Update(agendamento);
                SGID.SaveChanges();

                if(!string.IsNullOrEmpty(Obs) && !string.IsNullOrWhiteSpace(Obs))
                {
                    var Observacao = new ObsAgendamento { AgendamentoId = agendamento.Id, User = agendamento.UsuarioAlterar, Obs = Obs, DataCriacao = DateTime.Now };

                    SGID.ObsAgendamentos.Add(Observacao);
                    SGID.SaveChanges();
                }

                var AgendamentoProduto = SGID.ProdutosAgendamentos.Where(x => x.AgendamentoId == agendamento.Id).ToList();

                AgendamentoProduto.ForEach(produto =>
                {
                    var produtoUpdate = Produtos.FirstOrDefault(x => x.Item == produto.CodigoProduto);

                    if (produtoUpdate != null)
                    {

                        produto.Quantidade = produtoUpdate.Und;
                        produto.ValorUnitario = produtoUpdate.PrcUnid;
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
                        AgendamentoId = agendamento.Id,
                        CodigoProduto = produto.Item,
                        Quantidade = produto.Und,
                        ValorUnitario = produto.PrcUnid,
                        ValorTotal = produto.Und * produto.PrcUnid,
                        CodigoTabela = CodTabela
                    };

                    SGID.ProdutosAgendamentos.Add(ProdXAgenda);
                    SGID.SaveChanges();
                });

                var AgendamentoPatris = SGID.PatrimoniosAgendamentos.Where(x => x.AgendamentoId == agendamento.Id).ToList();

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
                        AgendamentoId = agendamento.Id,
                        Patrimonio = patri.Descri
                    };

                    SGID.PatrimoniosAgendamentos.Add(ProdXAgenda);
                    SGID.SaveChanges();
                });

                string Pasta = $"{_WEB.WebRootPath}/AnexosAgendamento";


                //Anexos
                #region Anexos
                if (!Directory.Exists(Pasta))
                {
                    Directory.CreateDirectory(Pasta);
                }

                if (Anexos01 != null)
                {
                    var anexo = SGID.AnexosAgendamentos.FirstOrDefault(x => x.AgendamentoId == agendamento.Id && x.NumeroAnexo == "01");

                    if (anexo == null)
                    {
                        var anexoAgenda = new AnexosAgendamentos
                        {
                            AgendamentoId = agendamento.Id,
                            AnexoCam = $"{agendamento.Id}01.{Anexos01.FileName.Split(".").Last()}",
                            NumeroAnexo = $"01"
                        };

                        string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            Anexos01.CopyTo(fileStream);
                        }

                        SGID.AnexosAgendamentos.Add(anexoAgenda);
                        SGID.SaveChanges();
                    }
                    else
                    {
                        anexo.AnexoCam = $"{agendamento.Id}01.{Anexos01.FileName.Split(".").Last()}";
                        string Caminho = $"{Pasta}/{anexo.AnexoCam}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            Anexos01.CopyTo(fileStream);
                        }

                        SGID.AnexosAgendamentos.Update(anexo);
                        SGID.SaveChanges();
                    }
                }
                if (Anexos02 != null)
                {
                    var anexo = SGID.AnexosAgendamentos.FirstOrDefault(x => x.AgendamentoId == agendamento.Id && x.NumeroAnexo == "02");

                    if (anexo == null)
                    {
                        var anexoAgenda = new AnexosAgendamentos
                        {
                            AgendamentoId = agendamento.Id,
                            AnexoCam = $"{agendamento.Id}02.{Anexos02.FileName.Split(".").Last()}",
                            NumeroAnexo = $"02"
                        };

                        string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            Anexos02.CopyTo(fileStream);
                        }

                        SGID.AnexosAgendamentos.Add(anexoAgenda);
                        SGID.SaveChanges();
                    }
                    else
                    {
                        anexo.AnexoCam = $"{agendamento.Id}02.{Anexos02.FileName.Split(".").Last()}";
                        string Caminho = $"{Pasta}/{anexo.AnexoCam}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            Anexos02.CopyTo(fileStream);
                        }

                        SGID.AnexosAgendamentos.Update(anexo);
                        SGID.SaveChanges();
                    }
                }
                if (Anexos03 != null)
                {
                    var anexo = SGID.AnexosAgendamentos.FirstOrDefault(x => x.AgendamentoId == agendamento.Id && x.NumeroAnexo == "03");

                    if (anexo == null)
                    {
                        var anexoAgenda = new AnexosAgendamentos
                        {
                            AgendamentoId = agendamento.Id,
                            AnexoCam = $"{agendamento.Id}03.{Anexos03.FileName.Split(".").Last()}",
                            NumeroAnexo = $"03"
                        };

                        string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            Anexos03.CopyTo(fileStream);
                        }

                        SGID.AnexosAgendamentos.Add(anexoAgenda);
                        SGID.SaveChanges();
                    }
                    else
                    {
                        anexo.AnexoCam = $"{agendamento.Id}03.{Anexos03.FileName.Split(".").Last()}";
                        string Caminho = $"{Pasta}/{anexo.AnexoCam}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            Anexos03.CopyTo(fileStream);
                        }

                        SGID.AnexosAgendamentos.Update(anexo);
                        SGID.SaveChanges();
                    }
                }
                if (Anexos04 != null)
                {
                    var anexo = SGID.AnexosAgendamentos.FirstOrDefault(x => x.AgendamentoId == agendamento.Id && x.NumeroAnexo == "04");
                    if (anexo == null)
                    {
                        var anexoAgenda = new AnexosAgendamentos
                        {
                            AgendamentoId = agendamento.Id,
                            AnexoCam = $"{agendamento.Id}04.{Anexos04.FileName.Split(".").Last()}",
                            NumeroAnexo = $"04"
                        };

                        string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            Anexos04.CopyTo(fileStream);
                        }

                        SGID.AnexosAgendamentos.Add(anexoAgenda);
                        SGID.SaveChanges();
                    }
                    else
                    {
                        anexo.AnexoCam = $"{agendamento.Id}04.{Anexos04.FileName.Split(".").Last()}";
                        string Caminho = $"{Pasta}/{anexo.AnexoCam}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            Anexos04.CopyTo(fileStream);
                        }

                        SGID.AnexosAgendamentos.Update(anexo);
                        SGID.SaveChanges();
                    }
                }
                if (Anexos05 != null)
                {
                    var anexo = SGID.AnexosAgendamentos.FirstOrDefault(x => x.AgendamentoId == agendamento.Id && x.NumeroAnexo == "05");

                    if(anexo == null)
                    {
                        var anexoAgenda = new AnexosAgendamentos
                        {
                            AgendamentoId = agendamento.Id,
                            AnexoCam = $"{agendamento.Id}05.{Anexos05.FileName.Split(".").Last()}",
                            NumeroAnexo = $"05"
                        };

                        string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            Anexos05.CopyTo(fileStream);
                        }

                        SGID.AnexosAgendamentos.Add(anexoAgenda);
                        SGID.SaveChanges();
                    }
                    else
                    {
                        anexo.AnexoCam = $"{agendamento.Id}05.{Anexos05.FileName.Split(".").Last()}";
                        string Caminho = $"{Pasta}/{anexo.AnexoCam}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            Anexos05.CopyTo(fileStream);
                        }

                        SGID.AnexosAgendamentos.Update(anexo);
                        SGID.SaveChanges();
                    }
                }

                #endregion

                /*if(User.IsInRole("Estoque"))
                {
                    if(agendamento.Empresa == "01")
                    {
                        return LocalRedirect("/logistica/ListarLogistica/01/1");  
                    }
                    else
                    {
                        return LocalRedirect("/logistica/ListarLogistica/03/1");
                    }
                }*/

                return LocalRedirect("/dashboards/dashboard/3");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "EditarAgendamento", user);

                return LocalRedirect("/error");
            }
        }
    }
}
