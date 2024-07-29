using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;
using SGID.Models.Email;
using System.Net.Mail;

namespace SGID.Pages.Formularios.Estoque
{
    public class ListarOcorrenciasModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public List<Ocorrencia> Ocorrencias { get; set; }
        
        private readonly IWebHostEnvironment _WEB;

        public ListarOcorrenciasModel(ApplicationDbContext sgid,TOTVSDENUOContext denuo,TOTVSINTERContext inter,IWebHostEnvironment web)
        {
            SGID = sgid;
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
            _WEB = web;
        }

        public void OnGet()
        {
            Ocorrencias = SGID.Ocorrencias.OrderByDescending(x=> x.Id).ToList();
        }

        public JsonResult OnGetEnviar(int Id)
        {
            try
            {
                var Ocorrencia = SGID.Ocorrencias.FirstOrDefault(x => x.Id == Id);

                var DataCirurgia = "";
                var DataOcorrencia = "";

                if (Ocorrencia.DataOcorrencia != null)
                {
                    DataOcorrencia = Ocorrencia.DataOcorrencia.Value.ToString("dd/MM/yyyy");
                }

                if (Ocorrencia.Cirurgia != null)
                {
                    DataCirurgia = Ocorrencia.Cirurgia.Value.ToString("dd/MM/yyyy");
                }

                var template = new
                {
                    Titulo = "Ocorrência",
                    Titulo2 = $"Nº {Id}",

                    DataOcorrencia = DataOcorrencia,
                    DescricaoOcorrencia = Ocorrencia.Descricao,
                    TipoOcorrencia = Ocorrencia.Problema,

                    Agendamento = Ocorrencia.Agendamento,
                    Cirurgia = DataCirurgia,
                    Cliente = Ocorrencia.Cliente,
                    Hospital = Ocorrencia.Hospital,
                    Medico = Ocorrencia.Medico,
                    Paciente = Ocorrencia.Paciente,

                    Patrimonio = $"{Ocorrencia.Patrimonio} - {Ocorrencia.DescPatri}",
                    Produto = $"{Ocorrencia.Produto} - {Ocorrencia.Descricao}",
                    Quant = Ocorrencia.Quantidade,

                    Acao = Ocorrencia.Acao,
                    Procedente = Ocorrencia.Procedente,
                    Cobrado = Ocorrencia.Cobrado,
                    Reposto = Ocorrencia.Reposto,

                    Obs = Ocorrencia.Obs

                };

                var email = "";

                if (Ocorrencia.Empresa == "01")
                {
                    email = ProtheusInter.Sa3010s.FirstOrDefault(x => x.A3Nome == Ocorrencia.Vendedor)?.A3Email;

                    if (email == null)
                    {
                        email = "ricardo.bassanese@intermedic.com.br";
                    }
                    else if (email.Contains(";"))
                    {
                        email = email.Split(";")[0];
                    }
                }
                else
                {
                    email = ProtheusDenuo.Sa3010s.FirstOrDefault(x => x.A3Nome == Ocorrencia.Vendedor)?.A3Email;

                    if (email == null)
                    {
                        email = "ricardo.bassanese@intermedic.com.br";
                    }
                    else if (email.Contains(";"))
                    {
                        email = email.Split(";")[0];
                    }

                }

                var mensagem = "";

                mensagem = EmailTemplate.LerArquivoHtml($"{_WEB.WebRootPath}/template/TemplateOcorrencia.html", template);

                SmtpClient client = new SmtpClient();
                client.Host = "smtp.office365.com";
                client.EnableSsl = true;
                client.Credentials = new System.Net.NetworkCredential("ti@intermedic.com.br", "interadm2018!*");
                MailMessage mail = new MailMessage();
                mail.Sender = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
                mail.From = new MailAddress("ti@intermedic.com.br", "ENVIADOR");
                mail.To.Add(new MailAddress(email, "RECEBEDOR"));

                if(Ocorrencia.Empresa == "01")
                {
                    mail.CC.Add(new MailAddress("estoque@intermedic.com.br", "estoque"));
                    mail.CC.Add(new MailAddress("qualidade@intermedic.com.br", "qualidade"));
                    mail.CC.Add(new MailAddress("comercial@intermedic.com.br", "comercial"));

                }
                else
                {
                    mail.CC.Add(new MailAddress("estoque@denuo.com.br", "Estoque Denuo"));
                    mail.CC.Add(new MailAddress("qualidade@denuo.com.br", "qualidade"));
                    mail.CC.Add(new MailAddress("comercial@denuo.com.br", "comercial"));
                }
                
                
                mail.CC.Add(new MailAddress("agendamento@intermedic.com.br", "Agendamento Intermedic"));
                mail.CC.Add(new MailAddress("admvendas@intermedic.com.br", "Adm Vendas"));
                
                mail.Subject = $"Ocorrência Nº {Id}";
                mail.Body = mensagem;
                mail.IsBodyHtml = true;
                mail.Priority = MailPriority.High;

                client.Send(mail);

                return new JsonResult("E-mail enviado com sucesso");

            }
            catch(Exception e)
            {
                return new JsonResult("Error: E-mail não enviado");
            }
        }
    }
}
