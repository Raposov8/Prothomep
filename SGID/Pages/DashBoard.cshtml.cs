using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.DTO;
using SGID.Models.Denuo;
using SGID.Models.Inter;
using SGID.Data.Models;
using SGID.Models.Email;
using System.Net.Mail;

namespace SGID.Pages
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

        public DashBoardModel(ILogger<DashBoardModel> logger,ApplicationDbContext sgid,TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter, IWebHostEnvironment wEB)
        {
            _logger = logger;
            SGID = sgid;
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
            _WEB = wEB;
        }
        public IActionResult OnGet(int id)
        {
           
            if (User.IsInRole("Instrumentador"))
            {
                return LocalRedirect("/instrumentador/dashboardinstrumentador");
            }
            else if (User.IsInRole("Diretoria"))
            {
                return LocalRedirect("/dashboards/DashBoardMetas");
            }
            else if (User.IsInRole("Financeiro") || User.IsInRole("Diretoria"))
            {
                return LocalRedirect("/relatorios/vexpenses/relatoriodespesas");
            }
            else if (User.IsInRole("SubDistribuidor"))
            {
                return LocalRedirect("/dashboards/DashBoardSubDistribuidor");
            }
            else if (User.IsInRole("Admin") || User.IsInRole("GestorVenda") || User.IsInRole("Venda") || User.IsInRole("Diretoria"))
            {
                Total = SGID.Agendamentos.Count();

                PendenteComercial = SGID.Agendamentos.Where(x => (x.StatusCotacao == 0 && x.StatusPedido == 1) || x.StatusPedido == 2).Count();

                RetornoCliente = SGID.Agendamentos.Where(x => x.StatusPedido == 6).Count();

                Respondidas = SGID.Agendamentos.Where(x => x.StatusPedido == 5).Count();

                NRespondidas = SGID.Agendamentos.Where(x => x.StatusPedido == 0).Count();

                Aprovadas = SGID.Agendamentos.Where(x=> x.StatusPedido == 3).Count();

                Perdidas = SGID.Agendamentos.Where(x => x.StatusPedido == 4).Count();

                Emergencia = SGID.Agendamentos.Where(x => x.Tipo == 1).Count();

                Agendamentos = id switch
                {
                    //Pendente Comercial Ida e volta da cotação
                    1 => SGID.Agendamentos.Where(x => (x.StatusCotacao == 0 && x.StatusPedido == 1) || x.StatusPedido == 2).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //Retorno Cliente
                    2 => SGID.Agendamentos.Where(x => x.StatusPedido == 6).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //NRespondidas
                    3 => SGID.Agendamentos.Where(x => x.StatusPedido == 0).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //Aprovadas
                    4 => SGID.Agendamentos.Where(x => x.StatusPedido == 3).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //Perdidas
                    5 => SGID.Agendamentos.Where(x => x.StatusPedido == 4).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //Emergencia
                    6 => SGID.Agendamentos.Where(x => x.Tipo == 1).OrderByDescending(x => x.DataCirurgia).ToList(),
                    //Respondidas
                    7 => SGID.Agendamentos.Where(x => x.StatusPedido == 5).OrderByDescending(x => x.DataCirurgia).ToList(),
                    _ => SGID.Agendamentos.OrderByDescending(x => x.DataCirurgia).ToList(),
                };

                Rejeicoes = SGID.RejeicaoMotivos.ToList();

                return Page();

            }
            else if(User.IsInRole("GestorComercial"))
            {
                if (User.Identity.Name.Split("@")[1].ToUpper()=="INTERMEDIC.COM.BR")
                {
                    return LocalRedirect("/dashboards/DashBoardGestorComercial/01");
                }
                else
                {
                    return LocalRedirect("/dashboards/DashBoardGestorComercial/03");
                }       
                
            }
            else if (User.IsInRole("Comercial"))
            {
                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {
                    return LocalRedirect("/dashboards/DashBoardComercial/01");
                }
                else
                {
                    return LocalRedirect("/dashboards/DashBoardComercial/03");
                }

            }
            else if (User.IsInRole("RH"))
            {
                return LocalRedirect("/rh/listargestaopessoal");
            }
            else if (User.IsInRole("Patrimonio"))
            {
                return LocalRedirect("/formularios/patrimonio/relatoriopatrimonio/01");
            }
            else
            {
                return LocalRedirect("/relatorios/controladoria/conserto");
            }
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
            var agendamento = SGID.Agendamentos.Select(x=> new
            {
                x.Cliente,
                x.CondPag,
                x.Convenio,
                DataAutorizacao = x.DataAutorizacao.ToString("dd/MM/yyyy"),
                DataCirurgia = x.DataCirurgia == null ? "SEM REGISTRO": x.DataCirurgia.Value.ToString("dd/MM/yyyy"),
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
            var anexos = SGID.AnexosAgendamentos.Where(x=>x.AgendamentoId == agenda).ToList();

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

        public JsonResult OnPostRecusarPedido(int AgendamentoId,string MotivoReij, string ObsReij)
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
            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

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
                    //Intermedic
                    var Codigo = ProtheusInter.Sa1010s.FirstOrDefault(x => x.A1Nome == AgendamentoAntigo.Medico && x.A1Clinter == "M" && x.DELET != "*" && x.A1Msblql != "1")?.A1Vend;

                    agendamento.VendedorLogin = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Cod == Codigo).A3Xlogin;
                }
                else
                {
                    //Denuo
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
                        Credentials = new System.Net.NetworkCredential("ti@intermedic.com.br", "interadm2018!*")
                    };
                    MailMessage mail = new MailMessage();
                    mail.Sender = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
                    mail.From = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
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
