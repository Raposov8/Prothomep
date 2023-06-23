using DocumentFormat.OpenXml.Office2010.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Cirurgias;
using SGID.Models.Denuo;
using SGID.Models.DTO;
using SGID.Models.Email;
using SGID.Models.Inter;
using System.Net.Mail;

namespace SGID.Pages.Cirurgias
{
    [Authorize(Roles = "Admin,GestorVenda,Venda,Comercial,GestorComercial")]
    public class NovoAgendamentoModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }

        private readonly IWebHostEnvironment _WEB;

        public NovoAgendamento Novo { get; set; } = new NovoAgendamento();
        
        public NovoAgendamentoModel(ApplicationDbContext sgid, TOTVSINTERContext protheus, TOTVSDENUOContext denuo, IWebHostEnvironment wEB)
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
                    Clientes = ProtheusInter.Sa1010s.Where(x => x.DELET != "*"  && x.A1Msblql != "1" && (x.A1Clinter == "C" || x.A1Clinter == "H" || x.A1Clinter == "M")).OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                    Convenio = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "C" && x.A1Msblql != "1").OrderBy(x => x.A1Nome).Select(x=>x.A1Nome).ToList(),
                    Medico = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Crm != "" ).OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                    Intrumentador = ProtheusInter.Pah010s.Where(x => x.DELET != "*" && x.PahMsblql != "1").OrderBy(x => x.PahNome).Select(x => x.PahNome).ToList(),
                    Hospital = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "H" && x.A1Msblql != "1").OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                    Procedimentos = SGID.Procedimentos.Where(x => x.Bloqueado == 0 && x.Empresa == "01").ToList(),
                    Patrimonios = ProtheusInter.Pa1010s.Where(x=> x.DELET != "*" && x.Pa1Msblql != "1").Select(x=> x.Pa1Despat).Distinct().ToList()
                };
            }
            else
            {
                //Denuo
                Novo = new NovoAgendamento 
                {
                    Clientes = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*"  && x.A1Msblql != "1" && (x.A1Clinter == "C" || x.A1Clinter == "H" || x.A1Clinter == "M")).OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                    Convenio = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "C" && x.A1Msblql != "1").OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                    Medico = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Crm != "").OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                    Intrumentador = ProtheusDenuo.Pah010s.Where(x => x.DELET != "*" && x.PahMsblql != "1").OrderBy(x => x.PahNome).Select(x => x.PahNome).ToList(),
                    Hospital = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "H" && x.A1Msblql != "1").OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
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
                                   ).Distinct().ToList()
                };
            }
        }

        public IActionResult OnPostAsync(string Empresa,string Cliente,string CondPag, string Tabela,string Vendedor,
            string Medico,string Matricula, string Paciente,string Convenio,string Instrumentador,string Hospital,
            DateTime? DataAgendamento,DateTime DataAutorizacao,string NumAutorizacao,string Senha,int Tipo,int Autorizado, string Obs,
            List<Produto> Produtos,List<Patrimonio> Patris, IFormCollection Anexos,string CodHospital,string CodConvenio,string CodTabela
            ,string CodCondPag,string Procedimento,string Indicacao)
        {
            try
            {
                var agendamento = new Agendamentos
                {
                    DataCriacao = DateTime.Now,
                    DataAlteracao = DateTime.Now,
                    Empresa = Empresa,
                    CodHospital = CodHospital,
                    Hospital = Hospital,
                    Cliente = Cliente,
                    CodCondPag = CodCondPag,
                    CondPag = CondPag,
                    CodTabela = CodTabela,
                    Tabela = Tabela,
                    Vendedor = Vendedor,
                    Medico = Medico,
                    Matricula = Matricula,
                    Paciente = Paciente,
                    CodConvenio = CodConvenio,
                    Convenio = Convenio,
                    Instrumentador = Instrumentador,
                    DataCirurgia = DataAgendamento,
                    DataAutorizacao = DataAutorizacao,
                    NumAutorizacao = NumAutorizacao,
                    Senha = Senha,
                    Tipo = Tipo,
                    Procedimento = Procedimento,
                    Autorizado = Autorizado,
                    Indicacao = Indicacao,
                    UsuarioCriacao = User.Identity.Name.Split("@")[0].ToUpper()
                };

                if (Empresa == "01")
                {
                    
                    //Intermedic
                    var Codigo = ProtheusInter.Sa1010s.FirstOrDefault(x => x.A1Nome == Medico && x.A1Clinter == "M" && x.DELET != "*" && x.A1Msblql != "1")?.A1Vend;

                    agendamento.VendedorLogin = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Cod == Codigo).A3Xlogin;
                }
                else
                {
                    //Denuo
                    var Codigo = ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.A1Nome == Medico && x.A1Clinter == "M" && x.DELET != "*" && x.A1Msblql != "1")?.A1Vend;

                    agendamento.VendedorLogin = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Cod == Codigo).A3Xlogin;

                }


                SGID.Agendamentos.Add(agendamento);
                SGID.SaveChanges();

                if (!string.IsNullOrEmpty(Obs) && !string.IsNullOrWhiteSpace(Obs))
                {
                    var Observacao = new ObsAgendamento { AgendamentoId = agendamento.Id, User = agendamento.UsuarioCriacao, Obs = Obs, DataCriacao = DateTime.Now };

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
                        ValorTotal = produto.VlrTotal,
                        CodigoTabela = CodTabela,
                    };

                    SGID.ProdutosAgendamentos.Add(ProdXAgenda);
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

                try
                {
                    #region Email

                    string VendedorEmail = "";
                    if (Empresa == "01")
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
                        Titulo = $"Novo Orçamento Nº {agendamento.Id} - {Paciente}"
                    };

                    mensagem = EmailTemplate.LerArquivoHtml($"{_WEB.WebRootPath}/template/TemplateCotacao.html", template);

                    SmtpClient client = new SmtpClient();
                    client.Host = "smtp.office365.com";
                    client.EnableSsl = true;
                    client.Credentials = new System.Net.NetworkCredential("ti@intermedic.com.br", "interadm2018!*");
                    MailMessage mail = new MailMessage();
                    mail.Sender = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
                    mail.From = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
                    mail.To.Add(new MailAddress($"{VendedorEmail}", "RECEBEDOR"));
                    mail.Subject = $"Novo Orçamento Nº {agendamento.Id} - {Paciente}";
                    mail.Body = mensagem;
                    mail.IsBodyHtml = true;
                    mail.Priority = MailPriority.High;

                    client.Send(mail);

                    #endregion
                }
                catch(Exception e)
                {
                    string user = User.Identity.Name.Split("@")[0].ToUpper();
                    Logger.Log(e, SGID, "NovoAgendamento", user);
                }
                return LocalRedirect("/dashboard/0");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento",user);

                return LocalRedirect("/error");
            }
        }

        public JsonResult OnPostAdicionarProd(string Codigo,string Empresa,string CodTab)
        {
            try
            {
                if (Empresa == "01") {

                        //Intermedic
                        var produto = (from SB10 in ProtheusInter.Sb1010s
                                       where SB10.B1Cod == Codigo.ToUpper() && SB10.B1Msblql != "1"
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


                        if (produto != null)
                        {
                            var preco = (from DA10 in ProtheusInter.Da1010s
                                         where DA10.DELET != "*" && DA10.Da1Codtab == CodTab && DA10.Da1Codpro == Codigo.ToUpper()
                                         select DA10.Da1Prcven).FirstOrDefault();
                            
                                var ViewProduto = new Produto
                                {
                                    Item = produto.B1Cod,
                                    Licit = produto.B1Solicit,
                                    Produtos = produto.B1Desc,
                                    Tuss = produto.B1Xtuss,
                                    Anvisa = produto.B1Reganvi,
                                    Marca = produto.B1Fabric,
                                    Und = 1,
                                    PrcUnid = preco,
                                    SegUnd = produto.B1Um,
                                    VlrTotal = preco,
                                };

                                return new JsonResult(ViewProduto);
                            
                        }

                }
                else
                {
                    //Denuo
                        var produto = (from SB10 in ProtheusDenuo.Sb1010s
                                       where SB10.B1Cod == Codigo.ToUpper() && SB10.B1Msblql != "1"
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



                        if (produto != null)
                        {
                            var preco = (from DA10 in ProtheusDenuo.Da1010s
                                         where DA10.DELET != "*" && DA10.Da1Codtab == CodTab && DA10.Da1Codpro == Codigo.ToUpper()
                                         select DA10.Da1Prcven).FirstOrDefault();

                       
                                var ViewProduto = new Produto
                                {
                                    Item = produto.B1Cod,
                                    Licit = produto.B1Solicit,
                                    Produtos = produto.B1Desc,
                                    Tuss = produto.B1Xtuss,
                                    Anvisa = produto.B1Reganvi,
                                    Marca = produto.B1Fabric,
                                    Und = 1,
                                    PrcUnid = preco,
                                    SegUnd = produto.B1Um,
                                    VlrTotal = preco,
                                    TipoOp = produto.B1Tipo,
                                };

                                return new JsonResult(ViewProduto);
                            
                        }
                    
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento AdicionarPatri",user);
            }

            return new JsonResult("");
        }
        public JsonResult OnPostAdicionarPatri(string Empresa,string Codigo)
        {
            try
            {
                if (Empresa == "01")
                {
                    var teste = this.Request.QueryString.ToString().Substring(44).Replace("%20", " ").Replace("%C3%83", "Ã");
                    var ViewObject = (from PA10 in ProtheusInter.Pa1010s
                                      join PAC in ProtheusInter.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                      from c in sr.DefaultIfEmpty()
                                      join SA10 in ProtheusInter.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                                      from a in st.DefaultIfEmpty()
                                      where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                                      && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == teste
                                      select new
                                      {
                                          Descri = PA10.Pa1Despat,
                                          KitBas = PA10.Pa1Kitbas,
                                      }).FirstOrDefault();


                    return new JsonResult(ViewObject);
                }
                else 
                {

                    var teste = this.Request.QueryString.ToString().Substring(44).Replace("%20"," ").Replace("%C3%83", "Ã");
                    var ViewObject = (from PA10 in ProtheusDenuo.Pa1010s
                                      join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                      from c in sr.DefaultIfEmpty()
                                      join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                                      from a in st.DefaultIfEmpty()
                                      where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                                      && c.DELET != "*" && a.DELET != "*"
                                      && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                                      && PA10.Pa1Despat == teste
                                      select new
                                      {
                                          Descri = PA10.Pa1Despat,
                                          KitBas = PA10.Pa1Kitbas,
                                      }).FirstOrDefault();



                    return new JsonResult(ViewObject);
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento Adicionar", user);
            }

            return new JsonResult("");
        }
        public JsonResult OnPostCliente(string Codigo, string Empresa)
        {
            try
            {
                var data = Convert.ToInt32(DateTime.Now.ToString("yyyy/MM/dd").Replace("/", ""));
                if (Empresa == "01")
                {

                    //Intermedic
                    var Cliente = ProtheusInter.Sa1010s.Select(x => new
                    {
                        Codigo = x.A1Cod,
                        Nreduz = x.A1Nome,
                        Cond = x.A1Cond,
                        Descon = x.A1Desccon,
                        Tabela = x.A1Tabela,
                        x.DELET,
                        x.A1Msblql
                    }).FirstOrDefault(x => x.Nreduz == Codigo && x.DELET != "*" && x.A1Msblql != "1");

                    //&& (int)(object)x.Da0Datate >= data

                    var Tabela = ProtheusInter.Da0010s.FirstOrDefault(x => x.Da0Codtab == Cliente.Tabela && x.DELET != "*" )?.Da0Descri;

                    var TabelaCliente = new
                    {
                        Cliente.Codigo,
                        Cliente.Nreduz,
                        Cliente.Cond,
                        Cliente.Descon,
                        Cliente.Tabela,
                        CondTabela = Tabela
                    };

                    return new JsonResult(TabelaCliente);

                }
                else
                {
                    //Denuo

                    var Cliente = ProtheusDenuo.Sa1010s.Select(x => new
                    {
                        Codigo = x.A1Cod,
                        Nreduz = x.A1Nome,
                        Cond = x.A1Cond,
                        Descon = x.A1Desccon,
                        Tabela = x.A1Tabela,
                        x.DELET,
                        x.A1Msblql
                    }).FirstOrDefault(x => x.Nreduz == Codigo && x.DELET != "*" && x.A1Msblql != "1");

                    var Tabela = ProtheusDenuo.Da0010s.FirstOrDefault(x => x.Da0Codtab == Cliente.Tabela && x.DELET != "*" && (int)(object)x.Da0Datate >= data)?.Da0Descri;


                    var TabelaCliente = new
                    {
                        Cliente.Codigo,
                        Cliente.Nreduz,
                        Cliente.Cond,
                        Cliente.Descon,
                        Cliente.Tabela,
                        CondTabela = Tabela
                    };

                    return new JsonResult(TabelaCliente);
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento AdicionarCliente", user);
            }

            return new JsonResult("");
        }
        public JsonResult OnGetVendedor(string Codigo,string Empresa)
        {
            try
            {
                if (Empresa == "01")
                {
                    //Intermedic
                    var Cliente = ProtheusInter.Sa1010s.FirstOrDefault(x => x.A1Nreduz == Codigo && x.DELET != "*" && x.A1Msblql != "1")?.A1Nvend;


                    return new JsonResult(Cliente);

                }
                else
                {
                    //Denuo
                    var Cliente = ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.A1Nreduz == Codigo && x.DELET != "*" && x.A1Msblql != "1")?.A1Nvend;


                    return new JsonResult(Cliente);
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento Vendedor", user);
            }

            return new JsonResult("");
        }
        public JsonResult OnGetCondicoes(string Codigo,string Empresa)
        {
            try
            {
                if (Empresa == "01")
                {
                    //Intermedic
                    var Cliente = ProtheusInter.Se4010s.FirstOrDefault(x => x.DELET != "*" && x.E4Msblql != "1" && x.E4Codigo == Codigo).E4Descri;

                    return new JsonResult(Cliente);
                }
                else
                {
                    //Denuo
                    var Cliente = ProtheusDenuo.Se4010s.FirstOrDefault(x => x.DELET != "*" && x.E4Msblql != "1" && x.E4Codigo == Codigo).E4Descri;

                    return new JsonResult(Cliente);
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento CodConvenio", user);
            }
            return new JsonResult("");
        }
        public JsonResult OnGetTabela(string Codigo,string Empresa)
        {
            try
            {
                var data = Convert.ToInt32(DateTime.Now.ToString("yyyy/MM/dd").Replace("/", ""));
                if (Empresa == "01")
                {
                    //Intermedic
                    //&& (int)(object)x.Da0Datate >= data
                    var Cliente = ProtheusInter.Da0010s.FirstOrDefault(x => x.Da0Codtab == Codigo && x.DELET != "*" ).Da0Descri;

                    return new JsonResult(Cliente);
                }
                else
                {
                    //Denuo
                    //&& (int)(object)x.Da0Datate >= data
                    var Cliente = ProtheusDenuo.Da0010s.FirstOrDefault(x => x.Da0Codtab == Codigo && x.DELET != "*").Da0Descri;

                    return new JsonResult(Cliente);
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento Tabela", user);
            }
            return new JsonResult("");
        }
        public JsonResult OnGetHospital(string Codigo,string Empresa)
        {
            try
            {
                if (Empresa == "01")
                {
                    //Intermedic
                    var Cliente = ProtheusInter.Sa1010s.FirstOrDefault(x => x.A1Cod == Codigo && x.A1Clinter == "H" && x.DELET != "*" && x.A1Msblql != "1").A1Nome;

                    return new JsonResult(Cliente);

                }
                else
                {
                    //Denuo
                    var Cliente = ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.A1Cod == Codigo && x.A1Clinter == "H" && x.DELET != "*" && x.A1Msblql != "1").A1Nome;

                    return new JsonResult(Cliente);
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento Hospital", user);
            }

            return new JsonResult("");
        }
        public JsonResult OnGetCodHospital(string Codigo, string Empresa)
        {
            try
            {
                if (Empresa == "01")
                {
                    //Intermedic
                    var Cliente = ProtheusInter.Sa1010s.FirstOrDefault(x => x.A1Nome == Codigo && x.A1Clinter == "H" && x.DELET != "*" && x.A1Msblql != "1").A1Cod;

                    return new JsonResult(Cliente);

                }
                else
                {
                    //Denuo
                    var Cliente = ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.A1Nome == Codigo && x.A1Clinter == "H" && x.DELET != "*" && x.A1Msblql != "1").A1Cod;

                    return new JsonResult(Cliente);
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento CodHospital", user);
            }

            return new JsonResult("");
        }
        public JsonResult OnGetConvenio(string Codigo,string Empresa)
        {
            try
            {
                if (Empresa == "01")
                {
                    //Intermedic
                    var Cliente = ProtheusInter.Sa1010s.FirstOrDefault(x => x.A1Cod == Codigo && x.A1Clinter == "C" && x.DELET != "*" && x.A1Msblql != "1").A1Nome;

                    return new JsonResult(Cliente);
                }
                else
                {
                    //Denuo
                    var Cliente = ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.A1Cod == Codigo && x.A1Clinter == "C" && x.DELET != "*" && x.A1Msblql != "1").A1Nome;

                    return new JsonResult(Cliente);
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento Convenio", user);
            }
            return new JsonResult("");
        }
        public JsonResult OnGetCodConvenio(string Codigo, string Empresa)
        {
            try
            {
                if (Empresa == "01")
                {
                    //Intermedic
                    var Cliente = ProtheusInter.Sa1010s.FirstOrDefault(x => x.A1Nome == Codigo && x.A1Clinter == "C" && x.DELET != "*" && x.A1Msblql != "1").A1Cod;
                    
                    return new JsonResult(Cliente);
                }
                else
                {
                    //Denuo
                    var Cliente = ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.A1Nome == Codigo && x.A1Clinter == "C" && x.DELET != "*" && x.A1Msblql != "1").A1Cod;

                    return new JsonResult(Cliente);
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovoAgendamento CodConvenio", user);
            }
            return new JsonResult("");
        }
    }
}
