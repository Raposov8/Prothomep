using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;
using SGID.Data;
using System.Net.Mail;
using SGID.Data.Models;
using SGID.Models.Email;

namespace SGID.Pages.RH
{
    [Authorize(Roles = "Admin,RH,Diretoria")]
    public class EditarGestaoPessoalModel : PageModel
    {

        private ApplicationDbContext SGID { get; set; }

        public SolicitacaoAcesso Acesso { get; set; }

        private readonly IWebHostEnvironment _WEB;

        public EditarGestaoPessoalModel(ApplicationDbContext sgid, IWebHostEnvironment wEB)
        {
            SGID = sgid;
            _WEB = wEB;
        }
        public void OnGet(int id)
        {
            Acesso = SGID.SolicitacaoAcessos.FirstOrDefault(x => x.Id == id);
        }

        public IActionResult OnPostAsync(int Id,DateTime DataEvento, string Empresa, string Nome, string Tipo, string Email,
           string Cargo, string Ramal, string NomeSub, string CargoSub, string Contratacao, string Obs, params string[] SelectedRoles)
        {
            try
            {
                var solicita = SGID.SolicitacaoAcessos.FirstOrDefault(x => x.Id == Id);

                solicita.DataAlteracao = DateTime.Now;
                solicita.DataEvento = DataEvento;
                solicita.Empresa = Empresa;
                solicita.Nome = Nome;
                solicita.Cargo = Cargo;
                solicita.Tipo = Tipo;
                solicita.Email = Email;
                solicita.NomeSub = NomeSub;
                solicita.Contratacao = Contratacao;
                solicita.CargoSub = CargoSub;
                solicita.Ramal = Ramal;
                solicita.Obs = Obs;

                var mensagem = "";

                SelectedRoles.ToList().ForEach(Soli =>
                {
                    switch (Soli)
                    {
                        case "Protheus": solicita.Protheus = true; break;
                        case "Ramal": solicita.IsRamal = true; break;
                        case "Maquina": solicita.Maquina = true; break;
                        case "Celular": solicita.Celular = true; break;
                        case "Impressora": solicita.Impressora = true; break;
                    }

                    mensagem += $"{Soli} <br/> ";
                });

                SGID.SolicitacaoAcessos.Update(solicita);
                SGID.SaveChanges();

                mensagem += $"<br/> Data Edição: {solicita.DataCriacao}";

                var teste = solicita.Contratacao == "S" ? "SIM" : "NÃO";

                var aumento = $"Aumento de Quadro: {teste}";

                if (teste == "NÃO")
                {
                    aumento += $"<br/> Nome Antigo Colaborador: {solicita.NomeSub} <br/> Cargo Antigo Colaborador: {solicita.CargoSub}";
                }

                var template = new
                {
                    Titulo = "EDIT:Solicitação de",
                    Titulo2 = "Acesso",
                    Data = DataEvento.ToString("dd/MM/yyyy"),
                    Email = Email,
                    Cargo = Cargo,
                    Nome = Nome,
                    Solicitacoes = mensagem,
                    Obs = Obs,
                    Aumento = aumento
                };

                mensagem = EmailTemplate.LerArquivoHtml($"{_WEB.WebRootPath}/template/TemplateEmail.html", template);

                SmtpClient client = new SmtpClient();
                client.Host = "smtp.office365.com";
                client.EnableSsl = true;
                client.Credentials = new System.Net.NetworkCredential("ti@flowmeds.com.br", "interadm2018!*");
                MailMessage mail = new MailMessage();
                mail.Sender = new MailAddress("ti@flowmeds.com.br", "ENVIADOR");
                mail.From = new MailAddress("ti@flowmeds.com.br", "ENVIADOR");
                mail.To.Add(new MailAddress("acessos@flowmeds.com.br", "RECEBEDOR"));
                mail.CC.Add(new MailAddress("rh@flowmeds.com.br"));
                mail.Subject = $"Correção:Solicitação de Acesso";
                mail.Body = mensagem;
                mail.IsBodyHtml = true;
                mail.Priority = MailPriority.High;

                client.Send(mail);

            }
            catch (Exception erro)
            {
                var user = User.Identity.Name.Split("@")[0];
                Logger.Log(erro, SGID, "EditarGestaoPessoal", user);
            }

            return LocalRedirect("/rh/listargestaopessoal");
        }
    }
}
