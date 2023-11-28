using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Email;
using System.Net.Mail;
using System.ServiceModel.Channels;

namespace SGID.Pages.RH
{
    [Authorize(Roles = "Admin,RH,Diretoria")]
    public class GestaoPessoalModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        private readonly IWebHostEnvironment _WEB;
        public GestaoPessoalModel(ApplicationDbContext sgid,IWebHostEnvironment web) 
        { 
            SGID = sgid;
            _WEB = web;
        }
        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(DateTime DataEvento,string Empresa,string Nome,string Tipo,string Email,
           string Cargo, string Ramal,string NomeSub,string CargoSub,string Contratacao,string Obs,params string[] SelectedRoles)
        {
            try
            {
                var solicita = new SolicitacaoAcesso
                {
                    DataCriacao = DateTime.Now,
                    DataEvento = DataEvento,
                    Empresa = Empresa,
                    Nome = Nome,
                    Cargo = Cargo,
                    Tipo = Tipo,
                    Email = Email,
                    Ramal = Ramal,
                    Obs = Obs,
                    NomeSub = NomeSub,
                    CargoSub = CargoSub,
                    Contratacao = Contratacao,
                    Usuario = User.Identity.Name.Split("@")[0]
                };

                var mensagem = "";
                if (DataEvento.ToString("dd/MM/yyyy") == "01/01/0001") return LocalRedirect("/error");

                SelectedRoles.ToList().ForEach(Soli =>
                {
                    switch (Soli)
                    {
                        case "Protheus": solicita.Protheus = true; break;
                        case "Ramal": solicita.IsRamal = true; break;
                        case "Maquina": solicita.Maquina = true; break;
                        case "Celular": solicita.Celular = true; break;
                        case "Impressora":solicita.Impressora = true;break;
                    }

                    mensagem += $"{Soli} <br/> ";
                });

                SGID.SolicitacaoAcessos.Add(solicita);
                SGID.SaveChanges();

                mensagem += $"<br/> Data Solicitação: {solicita.DataCriacao:dd/MM/yyyy HH:mm:ss}";

                var template = new
                {
                    Titulo = "Solicitação de",
                    Titulo2 = "Acesso",
                    Data = DataEvento.ToString("dd/MM/yyyy"),
                    Email = Email,
                    Cargo = Cargo,
                    Nome = Nome,
                    Solicitacoes = mensagem,
                    Obs = Obs
                };

                mensagem = EmailTemplate.LerArquivoHtml($"{_WEB.WebRootPath}/template/TemplateEmail.html",template);

                SmtpClient client = new SmtpClient
                {
                    Host = "smtp.office365.com",
                    EnableSsl = true,
                    Credentials = new System.Net.NetworkCredential("ti@intermedic.com.br", "interadm2018!*")
                };
                MailMessage mail = new MailMessage();
                mail.Sender = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
                mail.From = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
                mail.To.Add(new MailAddress("acessos@intermedic.com.br", "RECEBEDOR"));
                mail.CC.Add(new MailAddress("rh@intermedic.com.br"));
                mail.Subject = "Solicitação de Acesso";
                mail.Body = mensagem;
                mail.IsBodyHtml = true;
                mail.Priority = MailPriority.Normal;
                
                client.Send(mail);
            
            }
            catch (Exception erro)
            {
                var user = User.Identity.Name.Split("@")[0];
                Logger.Log(erro, SGID, "GestaoPessoal", user);
             }

            return LocalRedirect("/rh/listargestaopessoal");
        }
    }
}
