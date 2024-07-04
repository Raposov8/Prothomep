using Intergracoes.Inpart.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;
using SGID.Data;
using SGID.Models.Cirurgias;
using SGID.Models.Denuo;
using SGID.Models.Inter;
using SGID.Models.DTO;
using SGID.Data.Models;

namespace SGID.Pages.Emergencia
{
    public class EmergenciaModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }

        private readonly IWebHostEnvironment _WEB;
        public NovoAgendamento Novo { get; set; } = new NovoAgendamento();
        public List<Produto> Produtos { get; set; } = new List<Produto>();

        public List<string> SearchProduto { get; set; } = new List<string>();
        public List<Cotacao> Cotacoes { get; set; }

        [BindProperty]
        public AgendamentoModel Agendamento { get; set; } = new AgendamentoModel();

        public EmergenciaModel(ApplicationDbContext sgid, TOTVSINTERContext protheus, TOTVSDENUOContext denuo, IWebHostEnvironment wEB)
        {
            SGID = sgid;
            ProtheusInter = protheus;
            ProtheusDenuo = denuo;
            _WEB = wEB;
        }

        public void OnGet(string id)
        {
            Novo = new NovoAgendamento();

            if (id == "01")
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

                SearchProduto = ProtheusInter.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1" && x.B1Tipo != "KT" && x.B1Comerci == "C").Select(x => x.B1Cod + "  " + x.B1Desc).Distinct().ToList();
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

                SearchProduto = ProtheusDenuo.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1" && x.B1Tipo != "KT").Select(x => x.B1Cod + "  " + x.B1Desc).Distinct().ToList();
            }
        }
        public IActionResult OnPostAsync(List<Produto> Produtos, List<Patrimonio> Patris, IFormCollection Anexos)
        {
            try
            {
                    var agendamento = new Agendamentos
                    {
                        DataCriacao = DateTime.Now,
                        DataAlteracao = DateTime.Now,
                        Empresa = Agendamento.Empresa,
                        CodHospital = Agendamento.CodHospital,
                        Hospital = Agendamento.Hospital,
                        Cliente = Agendamento.Cliente,
                        CodCondPag = Agendamento.CodCondPag,
                        CondPag = Agendamento.CondPag,
                        CodTabela = Agendamento.CodTabela,
                        Tabela = Agendamento.Tabela,
                        Vendedor = Agendamento.Vendedor,
                        Medico = Agendamento.Medico,
                        Matricula = Agendamento.Matricula,
                        Paciente = Agendamento.Paciente,
                        CodConvenio = Agendamento.CodConvenio,
                        Convenio = Agendamento.Convenio,
                        Instrumentador = Agendamento.Instrumentador,
                        DataCirurgia = Agendamento.DataAgendamento,
                        DataAutorizacao = Agendamento.DataAutorizacao,
                        NumAutorizacao = Agendamento.NumAutorizacao,
                        Senha = Agendamento.Senha,
                        Tipo = Agendamento.Tipo,
                        DataEntrega = Agendamento.DataEntrega,
                        Procedimento = Agendamento.Procedimento,
                        Autorizado = Agendamento.Autorizado,
                        Indicacao = Agendamento.Indicacao,
                        StatusPedido = 7,
                        StatusLogistica = 0,
                        UsuarioCriacao = User.Identity.Name.Split("@")[0].ToUpper()
                    };


                    if (Agendamento.Empresa == "01")
                    {
                        //Intermedic
                        agendamento.VendedorLogin = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Nreduz == Agendamento.Vendedor).A3Xlogin;

                    }
                    else
                    {
                        //Denuo
                        agendamento.VendedorLogin = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Nreduz == Agendamento.Vendedor).A3Xlogin;

                    }


                    SGID.Agendamentos.Add(agendamento);
                    SGID.SaveChanges();

                    if (!string.IsNullOrEmpty(Agendamento.Obs) && !string.IsNullOrWhiteSpace(Agendamento.Obs))
                    {
                        var Observacao = new ObsAgendamento { AgendamentoId = agendamento.Id, User = agendamento.UsuarioCriacao, Obs = Agendamento.Obs, DataCriacao = DateTime.Now };

                        SGID.ObsAgendamentos.Add(Observacao);
                        SGID.SaveChanges();
                    }

                    Produtos.ForEach(produto =>
                    {
                        var ProdXAgenda = new ProdutosAgendamentos
                        {
                            AgendamentoId = agendamento.Id,
                            CodigoProduto = produto.Item,
                            Quantidade = produto.Und,
                            ValorUnitario = produto.PrcUnid,
                            ValorTotal = produto.Und * produto.PrcUnid,
                            CodigoTabela = Agendamento.CodTabela,
                        };

                        var avulso = new AvulsosAgendamento
                        {
                            AgendamentoId = agendamento.Id,
                            Produto = produto.Item,
                            Quantidade = produto.Und
                        };

                        SGID.ProdutosAgendamentos.Add(ProdXAgenda);
                        SGID.AvulsosAgendamento.Add(avulso);
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

                    if (!Directory.Exists(Pasta))
                    {
                        Directory.CreateDirectory(Pasta);
                    }

                    //Anexos
                    #region Anexos
                    int i = 1;

                    foreach (var anexo in Anexos.Files)
                    {
                        var anexoAgenda = new AnexosAgendamentos
                        {
                            AgendamentoId = agendamento.Id,
                            AnexoCam = $"{agendamento.Id}0{i}.{anexo.FileName.Split(".").Last()}",
                            NumeroAnexo = $"0{i}"
                        };

                        string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";
                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            anexo.CopyTo(fileStream);
                        }

                        SGID.AnexosAgendamentos.Add(anexoAgenda);
                        SGID.SaveChanges();

                        i++;
                    }

                    #endregion

                    /*var link = $"https://gidd.com.br/cotacoes/confirmarcotacao/{agendamento.Id}";

                    try
                    {
                        #region Email

                        string VendedorEmail = "";
                        List<string> GestorEmail = new List<string>();
                        if (Agendamento.Empresa == "01")
                        {
                            var vendedor = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Xlogin == agendamento.VendedorLogin);

                            VendedorEmail = vendedor?.A3Email;
                            if (!vendedor.A3Xlogsup.IsNullOrEmpty())
                            {
                                GestorEmail.Add(vendedor.A3Xlogsup.ToLower().Trim() + "@intermedic.com.br");
                            }
                        }
                        else
                        {
                            var vendedor = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Xlogin == agendamento.VendedorLogin);
                            VendedorEmail = vendedor?.A3Email;
                            if (!vendedor.A3Xlogsup.IsNullOrEmpty())
                            {
                                GestorEmail.Add(vendedor.A3Xlogsup.ToLower().Trim() + "@denuo.com.br");
                            }

                            if (GestorEmail.Contains("leonardo.brito@denuo.com.br"))
                            {
                                GestorEmail.Add("artemio.costa@denuo.com.br");
                            }
                        }

                        if (VendedorEmail.Contains(';'))
                        {
                            VendedorEmail = VendedorEmail.Split(";")[0];
                        }

                        var template = new
                        {
                            Titulo = $"{agendamento.Id} - {Agendamento.Paciente}",
                            DataCirurgia = agendamento.DataCirurgia.HasValue ? agendamento.DataCirurgia.Value.ToString("dd/MM/yyyy HH:mm") : "Sem Data",
                            Paciente = agendamento.Paciente,
                            Hospital = agendamento.Hospital,
                            Cliente = agendamento.Cliente,
                            Medico = agendamento.Medico,
                            Obs = !string.IsNullOrEmpty(Agendamento.Obs) && !string.IsNullOrWhiteSpace(Agendamento.Obs) ? Agendamento.Obs : "Sem Observação",
                            Link = link
                        };

                        var mensagem = EmailTemplate.LerArquivoHtml($"{_WEB.WebRootPath}/template/TemplateCotacao.html", template);

                        SmtpClient client = new SmtpClient();
                        client.Host = "smtp.office365.com";
                        client.EnableSsl = true;
                        client.Credentials = new System.Net.NetworkCredential("ti@intermedic.com.br", "interadm2018!*");
                        MailMessage mail = new MailMessage();
                        mail.Sender = new MailAddress("ti@intermedic.com.br", "GID");
                        mail.From = new MailAddress("ti@intermedic.com.br", "GID");
                        mail.To.Add(new MailAddress($"{VendedorEmail}", "RECEBEDOR"));

                        GestorEmail.Where(x => !x.IsNullOrEmpty()).ToList().ForEach(x =>
                        {
                            mail.CC.Add(new MailAddress($"{x}", "COPIA"));
                        });
                        mail.Bcc.Add(new MailAddress("andre.souza@intermedic.com.br", "ANDRE SOUZA"));
                        mail.Subject = $"Emergencia Nº {agendamento.Id} - {Agendamento.Paciente}";
                        mail.Body = mensagem;
                        mail.IsBodyHtml = true;
                        mail.Priority = MailPriority.High;

                        client.Send(mail);

                        #endregion
                    }
                    catch (Exception e)
                    {
                        string user = User.Identity.Name.Split("@")[0].ToUpper();
                        Logger.Log(e, SGID, "Emergencia Email", user);
                    }*/

                    return LocalRedirect("/dashboards/dashboard/0");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Emergencia", user);

                return LocalRedirect("/error");
            }
        }
        public class AgendamentoModel
        {
            public string? Empresa { get; set; }
            public string? Cliente { get; set; }
            public string? CondPag { get; set; }
            public string? Tabela { get; set; }
            public string? Vendedor { get; set; }
            public string? Medico { get; set; }
            public string? Matricula { get; set; }
            public string? Paciente { get; set; }
            public string? Convenio { get; set; }
            public string? Instrumentador { get; set; }
            public string? Hospital { get; set; }
            public DateTime? DataAgendamento { get; set; }
            public DateTime? DataAutorizacao { get; set; }
            public DateTime? DataEntrega { get; set; }
            public string? NumAutorizacao { get; set; }
            public string? Senha { get; set; }
            public int Tipo { get; set; }
            public int Autorizado { get; set; }
            public string? Obs { get; set; }
            public string? CodHospital { get; set; }
            public string? CodConvenio { get; set; }
            public string? CodTabela { get; set; }
            public string? CodCondPag { get; set; }
            public string? Procedimento { get; set; }
            public string? Indicacao { get; set; }
        }
    }
}
