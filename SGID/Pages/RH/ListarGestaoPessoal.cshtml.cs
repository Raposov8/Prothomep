using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Email;
using System.Net.Mail;

namespace SGID.Pages.RH
{
    [Authorize (Roles = "Admin,RH,Diretoria")]
    public class ListarGestaoPessoalModel : PageModel
    {

        private ApplicationDbContext SGID { get; set; }
        private readonly IWebHostEnvironment _WEB;
        public List<SolicitacaoAcesso> Acessos { get; set; }
        

        public ListarGestaoPessoalModel(ApplicationDbContext sgid,IWebHostEnvironment web) 
        {
            SGID = sgid;
            _WEB = web;
        }  
        public void OnGet()
        {
            Acessos = SGID.SolicitacaoAcessos.OrderByDescending(x=> x.DataCriacao).ToList();
        }

        public IActionResult OnGetCancelar(int Id)
        {
            var solicita = SGID.SolicitacaoAcessos.FirstOrDefault(x => x.Id == Id);

            var mensagem = "";

            var template = new
            {
                Titulo = "Cancelamento de ",
                Titulo2 = "Solicitação de Acesso",
                Nome = solicita.Nome,
                Email = solicita.Email,
            };

            mensagem = EmailTemplate.LerArquivoHtml($"{_WEB.WebRootPath}/template/TemplateCancelaSolicitacao.html", template);

            SmtpClient client = new SmtpClient();
            client.Host = "smtp.office365.com";
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential("ti@intermedic.com.br", "interadm2018!*");
            MailMessage mail = new MailMessage();
            mail.Sender = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
            mail.From = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
            mail.To.Add(new MailAddress("acessos@intermedic.com.br", "RECEBEDOR"));
            mail.CC.Add(new MailAddress("rh@intermedic.com.br"));
            mail.Subject = "Cancelamento de Solicitação de Acesso";
            mail.Body = mensagem;
            mail.IsBodyHtml = true;
            mail.Priority = MailPriority.High;

            client.Send(mail);

            SGID.SolicitacaoAcessos.Remove(solicita);
            SGID.SaveChanges();

            Acessos = SGID.SolicitacaoAcessos.Where(x => x.DataEvento >= DateTime.Now).ToList();

            return Page();
        }

        public IActionResult OnPostTermo(int Id,IFormCollection Anexos)
        {
            string Pasta = $"{_WEB.WebRootPath}/AcessoTermos";

            if (!Directory.Exists(Pasta))
            {
                Directory.CreateDirectory(Pasta);
            }

            //Anexos
            #region Anexos

            foreach (var anexo in Anexos.Files)
            {
                var anexoTermo = new AcessoTermo
                {
                    AcessoId = Id,
                    Caminho = $"{Id}.{anexo.FileName.Split(".").Last()}",
                    Nome = anexo.FileName
                };

                string Caminho = $"{Pasta}/{anexoTermo.Caminho}";
                using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                {
                    anexo.CopyTo(fileStream);
                }

                SGID.AcessoTermos.Add(anexoTermo);
                SGID.SaveChanges();
            }

            #endregion

            return LocalRedirect("/RH/ListarGestaoPessoal");

        }
    }
}
