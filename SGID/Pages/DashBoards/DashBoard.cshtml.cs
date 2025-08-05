using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;
using SGID.Data;
using SGID.Models.Denuo;
using SGID.Models.Email;
using SGID.Models.Inter;
using System.Net.Mail;
using SGID.Data.Models;
using Microsoft.AspNetCore.Authorization;
using SGID.Models.DTO;
using Microsoft.IdentityModel.Tokens;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace SGID.Pages.DashBoards
{
    [Authorize]
    public class DashBoardModel : PageModel
    {
        private readonly ILogger<DashBoardModel> _logger;
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        public List<Agendamentos> Agendamentos { get; set; }
        public List<RejeicaoMotivos> Rejeicoes { get; set; }

        private readonly IWebHostEnvironment _WEB;

        public int Total { get; set; }
        public int PendenteComercial { get; set; }
        public int RetornoCliente { get; set; }
        public int Respondidas { get; set; }
        public int NRespondidas { get; set; }
        public int Aprovadas { get; set; }
        public int Perdidas { get; set; }
        public int Emergencia { get; set; }
        public int Permanente { get; set; }

        public DashBoardModel(ILogger<DashBoardModel> logger, ApplicationDbContext sgid, TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter, IWebHostEnvironment wEB)
        {
            _logger = logger;
            SGID = sgid;
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
            _WEB = wEB;
        }
        public IActionResult OnGet(int id)
        {
                Total = SGID.Agendamentos.Where(x=> x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).Count();

                PendenteComercial = SGID.Agendamentos.Where(x => ((x.StatusCotacao == 0 && x.StatusPedido == 1) || x.StatusPedido == 2) && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).Count();

                RetornoCliente = SGID.Agendamentos.Where(x => x.StatusPedido == 6 && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).Count();

                Respondidas = SGID.Agendamentos.Where(x => x.StatusPedido == 5 && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).Count();

                NRespondidas = SGID.Agendamentos.Where(x => x.StatusPedido == 0 && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).Count();

                Aprovadas = SGID.Agendamentos.Where(x => (x.StatusPedido == 3 || x.StatusPedido == 7) && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).Count();

                Perdidas = SGID.Agendamentos.Where(x => x.StatusPedido == 4 && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).Count();

                Emergencia = SGID.Agendamentos.Where(x => x.Tipo == 1 || x.Tipo == 2).Count();
                
                Permanente = SGID.Agendamentos.Where(x => x.Tipo == 4 && x.StatusLogistica != 6).Count();

            Agendamentos = id switch
                {
                    //Pendente Comercial Ida e volta da cotação
                    1 => SGID.Agendamentos.Where(x => ((x.StatusCotacao == 0 && x.StatusPedido == 1) || x.StatusPedido == 2) && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //Retorno Cliente
                    2 => SGID.Agendamentos.Where(x => x.StatusPedido == 6 && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //NRespondidas
                    3 => SGID.Agendamentos.Where(x => x.StatusPedido == 0 && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //Aprovadas
                    4 => SGID.Agendamentos.Where(x => (x.StatusPedido == 3 || x.StatusPedido == 7) && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //Perdidas
                    5 => SGID.Agendamentos.Where(x => x.StatusPedido == 4 && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //Emergencia
                    6 => SGID.Agendamentos.Where(x => x.Tipo == 1 || x.Tipo == 2).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //Respondidas
                    7 => SGID.Agendamentos.Where(x => x.StatusPedido == 5 && x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //Permanente
                    8 => SGID.Agendamentos.Where(x => x.Tipo == 4 && x.StatusLogistica != 6).OrderByDescending(x => x.DataCirurgia).ToList(),

                    _ => SGID.Agendamentos.Where(x => x.Tipo != 1 && x.Tipo != 2 && x.Tipo != 4).OrderByDescending(x => x.DataCirurgia).ToList(),
                };

                Rejeicoes = SGID.RejeicaoMotivos.ToList();

                return Page();
        }

        public JsonResult OnGetItens(int agenda)
        {
            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == agenda);

            var codigos = SGID.ProdutosAgendamentos.Where(x => x.AgendamentoId == agendamento.Id)
                    .Select(x => new
                    {
                        produto = x.CodigoProduto,
                        unidade = x.Quantidade,
                        valor = x.ValorTotal
                    }).ToList();


            var produtos = new List<Produto>();

            if (agendamento.Empresa == "01")
            {
                codigos.ForEach(x =>
                {
                    var produto = ProtheusInter.Sb1010s.Where(c => c.B1Cod == x.produto).Select(c => new Produto
                    {
                        Item = c.B1Cod,
                        Licit = c.B1Solicit,
                        Produtos = c.B1Desc,
                        Tuss = c.B1Xtuss,
                        Anvisa = c.B1Reganvi,
                        Marca = c.B1Fabric,
                        Und = x.unidade,
                        PrcUnid = c.B1Uprc,
                        SegUnd = c.B1Um,
                        VlrTotal = x.valor,
                        TipoOp = c.B1Tipo,
                    }).First();

                    produtos.Add(produto);
                });
            }
            else
            {
                codigos.ForEach(x =>
                {
                    var produto = ProtheusDenuo.Sb1010s.Where(c => c.B1Cod == x.produto).Select(c => new Produto
                    {
                        Item = c.B1Cod,
                        Licit = c.B1Solicit,
                        Produtos = c.B1Desc,
                        Tuss = c.B1Xtuss,
                        Anvisa = c.B1Reganvi,
                        Marca = c.B1Fabric,
                        Und = x.unidade,
                        PrcUnid = c.B1Uprc,
                        SegUnd = c.B1Um,
                        VlrTotal = x.valor,
                        TipoOp = c.B1Tipo,
                    }).First();

                    produtos.Add(produto);
                });
            }

            return new JsonResult(produtos);
        }

        public JsonResult OnGetDados(int agenda)
        {
            var agendamento = SGID.Agendamentos.Select(x => new
            {
                x.Cliente,
                x.CondPag,
                x.Convenio,
                DataAutorizacao = x.DataAutorizacao == null ? "SEM REGISTRO" : x.DataAutorizacao.Value.ToString("dd/MM/yyyy"),
                DataCirurgia = x.DataCirurgia == null ? "SEM REGISTRO" : x.DataCirurgia.Value.ToString("dd/MM/yyyy"),
                x.Hospital,
                x.Id,
                x.Instrumentador,
                x.Matricula,
                x.Medico,
                x.NumAutorizacao,
                x.Observacao,
                x.Paciente,
                x.Senha,
                x.Tabela,
                x.Tipo,
                x.Vendedor,
                x.Empresa,
                x.Autorizado,
                x.StatusPedido,
                DataCriacao = x.DataCriacao.ToString("dd/MM/yyyy")
            }).FirstOrDefault(x => x.Id == agenda);

            return new JsonResult(agendamento);
        }

        public JsonResult OnPostAnexos(int agenda)
        {
            var anexos = SGID.AnexosAgendamentos.Where(x => x.AgendamentoId == agenda).ToList();

            return new JsonResult(anexos);
        }

        public JsonResult OnPostStatus(int agenda)
        {
            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == agenda);

            agendamento.StatusPedido = 2;
            agendamento.StatusCotacao = 0;
            agendamento.DataAlteracao = DateTime.Now;

            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

            return new JsonResult("");
        }

        public JsonResult OnPostRecusarPedido(int AgendamentoId, string MotivoReij, string ObsReij)
        {

            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == AgendamentoId);

            agendamento.UsuarioRejeicao = User.Identity.Name.Split("@")[0].ToUpper();
            agendamento.StatusPedido = 4;
            agendamento.MotivoRejeicao = MotivoReij;
            agendamento.ObsRejeicao = ObsReij;
            agendamento.DataAlteracao = DateTime.Now;
            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

            return new JsonResult("");
        }

        public JsonResult OnGetRecusaPedido(int AgendamentoId)
        {
            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == AgendamentoId);

            var Visualiza = new
            {
                Motivo = agendamento.MotivoRejeicao,
                Obs = agendamento.ObsRejeicao
            };

            return new JsonResult(Visualiza);
        }

        public JsonResult OnPostAceitarPedido(int AgendamentoId)
        {
            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == AgendamentoId);

            agendamento.StatusPedido = 3;
            agendamento.DataAlteracao = DateTime.Now;
            agendamento.UsuarioAprova = User.Identity.Name.Split("@")[0].ToUpper();
            agendamento.DataAprova = DateTime.Now;
            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

            try
            {
                #region Email

                string VendedorEmail = "";
                List<string> GestorEmail = new List<string>();
                if (agendamento.Empresa == "01")
                {
                    var vendedor = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Xlogin == agendamento.VendedorLogin);

                    VendedorEmail = vendedor?.A3Email;
                    if (!vendedor.A3Xlogsup.IsNullOrEmpty())
                    {
                        GestorEmail.Add(vendedor.A3Xlogsup.ToLower().Trim() + "@flowmeds.com.br");
                    }
                }
                else
                {
                    var vendedor = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Xlogin == agendamento.VendedorLogin);
                    VendedorEmail = vendedor?.A3Email;
                    if (!vendedor.A3Xlogsup.IsNullOrEmpty())
                    {
                        GestorEmail.Add(vendedor.A3Xlogsup.ToLower().Trim() + "@flowmeds.com.br");
                    }

                    if (GestorEmail.Contains("leonardo.brito@flowmeds.com.br"))
                    {
                        GestorEmail.Add("artemio.costa@flowmeds.com.br");
                    }
                }

                if (VendedorEmail.Contains(';'))
                {
                    VendedorEmail = VendedorEmail.Split(";")[0];
                }

                var template = new
                {
                    Titulo = $"{agendamento.Id} - {agendamento.Paciente}",
                    Cotacao = agendamento.Id,
                    DataCirurgia = agendamento.DataCirurgia.HasValue ? agendamento.DataCirurgia.Value.ToString("dd/MM/yyyy HH:mm") : "Sem Data",
                    Paciente = agendamento.Paciente,
                    Hospital = agendamento.Hospital,
                    Medico = agendamento.Medico,
                    Link = $"https://gidd.com.br/cotacoes/confirmarcotacoes/{agendamento.Id}"
                };

                var mensagem = EmailTemplate.LerArquivoHtml($"{_WEB.WebRootPath}/template/TemplateAprovacao.html", template);

                SmtpClient client = new SmtpClient();
                client.Host = "smtp.office365.com";
                client.EnableSsl = true;
                client.Credentials = new System.Net.NetworkCredential("ti@flowmeds.com.br", "interadm2018!*");
                MailMessage mail = new MailMessage();
                mail.Sender = new MailAddress("ti@flowmeds.com.br", "GID");
                mail.From = new MailAddress("ti@flowmeds.com.br", "GID");
                mail.To.Add(new MailAddress($"{VendedorEmail}", "RECEBEDOR"));
                GestorEmail.Where(x => !x.IsNullOrEmpty()).ToList().ForEach(x =>
                {
                    mail.CC.Add(new MailAddress($"{x}", "COPIA"));
                });
                mail.Bcc.Add(new MailAddress("andre.souza@flowmeds.com.br", "ANDRE SOUZA"));
                mail.Subject = $"Orçamento Nº {agendamento.Id} - {agendamento.Paciente}";
                mail.Body = mensagem;
                mail.IsBodyHtml = true;
                mail.Priority = MailPriority.High;

                client.Send(mail);

                #endregion
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Dashboard Email", user);
            }

            return new JsonResult("");
        }

        public JsonResult OnPostDuplicarPedido(int AgendamentoId)
        {
            try
            {
                var AgendamentoAntigo = SGID.Agendamentos.FirstOrDefault(x => x.Id == AgendamentoId);

                var ProdutosAntigo = SGID.ProdutosAgendamentos.Where(x => x.AgendamentoId == AgendamentoId).ToList();

                var PatrimoniosAntigo = SGID.PatrimoniosAgendamentos.Where(x => x.AgendamentoId == AgendamentoId).ToList();

                var AvulsosAntigo = SGID.AvulsosAgendamento.Where(x => x.AgendamentoId == AgendamentoId).ToList();

                var AnexosAntigo = SGID.AnexosAgendamentos.Where(x => x.AgendamentoId == AgendamentoId).ToList();

                var agendamento = new Agendamentos
                {
                    DataCriacao = DateTime.Now,
                    DataAlteracao = DateTime.Now,
                    Empresa = AgendamentoAntigo.Empresa,
                    CodHospital = AgendamentoAntigo.CodHospital,
                    Hospital = AgendamentoAntigo.Hospital,
                    Cliente = AgendamentoAntigo.Cliente,
                    CodCondPag = AgendamentoAntigo.CodCondPag,
                    CondPag = AgendamentoAntigo.CondPag,
                    CodTabela = AgendamentoAntigo.CodTabela,
                    Tabela = AgendamentoAntigo.Tabela,
                    Vendedor = AgendamentoAntigo.Vendedor,
                    Medico = AgendamentoAntigo.Medico,
                    Matricula = AgendamentoAntigo.Matricula,
                    Paciente = AgendamentoAntigo.Paciente,
                    CodConvenio = AgendamentoAntigo.CodConvenio,
                    Convenio = AgendamentoAntigo.Convenio,
                    Instrumentador = AgendamentoAntigo.Instrumentador,
                    DataCirurgia = AgendamentoAntigo.DataCirurgia,
                    DataAutorizacao = AgendamentoAntigo.DataAutorizacao,
                    NumAutorizacao = AgendamentoAntigo.NumAutorizacao,
                    Senha = AgendamentoAntigo.Senha,
                    Tipo = AgendamentoAntigo.Tipo,
                    Procedimento = AgendamentoAntigo.Procedimento,
                    Autorizado = AgendamentoAntigo.Autorizado,
                    Indicacao = AgendamentoAntigo.Indicacao,
                    UsuarioCriacao = User.Identity.Name.Split("@")[0].ToUpper()
                };

                if (AgendamentoAntigo.Empresa == "01")
                {
                    //J&J
                    var Codigo = ProtheusInter.Sa1010s.FirstOrDefault(x => x.A1Nome == AgendamentoAntigo.Medico && x.A1Clinter == "M" && x.DELET != "*" && x.A1Msblql != "1")?.A1Vend;

                    agendamento.VendedorLogin = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Cod == Codigo).A3Xlogin;
                }
                else
                {
                      //FLOWMED
                    var Codigo = ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.A1Nome == AgendamentoAntigo.Medico && x.A1Clinter == "M" && x.DELET != "*" && x.A1Msblql != "1")?.A1Vend;

                    agendamento.VendedorLogin = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Cod == Codigo).A3Xlogin;
                }

                SGID.Agendamentos.Add(agendamento);
                SGID.SaveChanges();

                ProdutosAntigo.ForEach(produto =>
                {
                    var ProdXAgenda = new ProdutosAgendamentos
                    {
                        AgendamentoId = agendamento.Id,
                        CodigoProduto = produto.CodigoProduto,
                        Quantidade = produto.Quantidade,
                        ValorTotal = produto.ValorTotal,
                        CodigoTabela = produto.CodigoTabela,
                    };

                    SGID.ProdutosAgendamentos.Add(ProdXAgenda);
                    SGID.SaveChanges();
                });

                PatrimoniosAntigo.ForEach(patri =>
                {
                    var ProdXAgenda = new PatrimonioAgendamento
                    {
                        AgendamentoId = agendamento.Id,
                        Patrimonio = patri.Patrimonio
                    };

                    SGID.PatrimoniosAgendamentos.Add(ProdXAgenda);
                    SGID.SaveChanges();
                });

                AvulsosAntigo.ForEach(avulso =>
                {
                    var avulso2 = new AvulsosAgendamento
                    {
                        AgendamentoId = agendamento.Id,
                        Produto = avulso.Produto,
                        Quantidade = avulso.Quantidade
                    };

                    SGID.AvulsosAgendamento.Add(avulso2);
                    SGID.SaveChanges();
                });

                string Pasta = $"{_WEB.WebRootPath}/AnexosAgendamento";

                if (!Directory.Exists(Pasta))
                {
                    Directory.CreateDirectory(Pasta);
                }
                //Anexos
                #region Anexos

                foreach (var anexo in AnexosAntigo)
                {

                    string CaminhoAntigo = $"{Pasta}/{anexo.AnexoCam}";


                    var anexoAgenda = new AnexosAgendamentos
                    {
                        AgendamentoId = agendamento.Id,
                        AnexoCam = $"{agendamento.Id}{anexo.NumeroAnexo}.{anexo.AnexoCam.Split(".").Last()}",
                        NumeroAnexo = anexo.NumeroAnexo
                    };

                    string Caminho = $"{Pasta}/{anexoAgenda.AnexoCam}";

                    using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                    {

                        var teste = new FileStream(CaminhoAntigo, FileMode.Open);

                        teste.CopyTo(fileStream);
                    }

                    SGID.AnexosAgendamentos.Add(anexoAgenda);
                    SGID.SaveChanges();

                }

                #endregion

                try
                {
                    #region Email

                    string VendedorEmail = "";
                    if (agendamento.Empresa == "01")
                    {
                        VendedorEmail = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Xlogin == agendamento.VendedorLogin)?.A3Email;
                    }
                    else
                    {
                        VendedorEmail = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Xlogin == agendamento.VendedorLogin)?.A3Email;
                    }

                    if (VendedorEmail.Contains(';'))
                    {
                        VendedorEmail = VendedorEmail.Split(";")[0];
                    }

                    var mensagem = "";

                    mensagem += $"";

                    var template = new
                    {
                        Titulo = $"Novo Orçamento Nº {agendamento.Id} - {agendamento.Paciente}"
                    };

                    mensagem = EmailTemplate.LerArquivoHtml($"{_WEB.WebRootPath}/template/TemplateCotacao.html", template);

                    SmtpClient client = new SmtpClient
                    {
                        Host = "smtp.office365.com",
                        EnableSsl = true,
                        Credentials = new System.Net.NetworkCredential("ti@flowmeds.com.br", "interadm2018!*")
                    };
                    MailMessage mail = new MailMessage();
                    mail.Sender = new MailAddress("ti@flowmeds.com.br", "ENVIADOR");
                    mail.From = new MailAddress("ti@flowmeds.com.br", "ENVIADOR");
                    mail.To.Add(new MailAddress($"{VendedorEmail}", "RECEBEDOR"));
                    mail.Subject = $"Novo Orçamento Nº {agendamento.Id} - {agendamento.Paciente}";
                    mail.Body = mensagem;
                    mail.IsBodyHtml = true;
                    mail.Priority = MailPriority.High;

                    client.Send(mail);

                    #endregion
                }
                catch (Exception e)
                {
                    string user = User.Identity.Name.Split("@")[0].ToUpper();
                    Logger.Log(e, SGID, "NovoAgendamento", user);
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento", user);
            }

            return new JsonResult("");
        }
    }
}
