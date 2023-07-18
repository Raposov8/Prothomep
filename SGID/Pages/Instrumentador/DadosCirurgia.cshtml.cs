using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.DTO;

namespace SGID.Pages.Instrumentador
{
    [Authorize(Roles = "Admin,Instrumentador,Diretoria")]
    public class DadosCirurgiaModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        public Agendamentos Agendamento { get; set; }
        private readonly IWebHostEnvironment _WEB;

        public DadosCirurgiaModel(ApplicationDbContext sgid, IWebHostEnvironment wEB)
        {
            SGID = sgid;
            _WEB = wEB;
        }
        public void OnGet(int id)
        {
           Agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == id);
        }

        public IActionResult OnPostAsync(int Id,string codigo,string NomePaciente,string NomeMedico,string NomeCliente,int Status,
            DateTime DataCirurgia,string Hora,string Procedimento,string Obs, IFormCollection Anexos01, IFormCollection Anexos02, 
            IFormCollection Anexos03, IFormCollection Anexos04, IFormCollection Anexos05)
        {

            var dados = new DadosCirurgia
            {
                DataCriacao = DateTime.Now,
                Codigo = codigo,
                NomePaciente = NomePaciente,
                NomeMedico = NomeMedico,
                NomeCliente = NomeCliente,
                Status = Status,
                DataCirurgia = DataCirurgia,
                HoraCirurgia = Hora,
                ProcedimentosExec = Procedimento,
                ObsIntercorrencia = Obs,
                AgendamentoId = Id
            };

            string Pasta = $"{_WEB.WebRootPath}/AnexosDados";

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

            return LocalRedirect("/instrumentador/dashboardinstrumentador");
        }
    }
}
