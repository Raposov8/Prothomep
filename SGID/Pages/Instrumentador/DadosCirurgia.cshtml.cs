using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;
using SGID.Models.DTO;
using SGID.Models.Inter;

namespace SGID.Pages.Instrumentador
{
    [Authorize]
    public class DadosCirurgiaModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext INTER { get; set; }
        private TOTVSDENUOContext DENUO { get; set; }
        public Agendamentos Agendamento { get; set; }
        private readonly IWebHostEnvironment _WEB;

        public List<string> SearchProduto { get; set; } = new List<string>();

        public DadosCirurgiaModel(ApplicationDbContext sgid, IWebHostEnvironment wEB,TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            SGID = sgid;
            _WEB = wEB;
            INTER = inter;
            DENUO = denuo;
        }

        public void OnGet(int id)
        {
           Agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == id);


            if(Agendamento.Empresa == "01") SearchProduto = INTER.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1" && x.B1Tipo != "KT" && x.B1Comerci=="C").Select(x => x.B1Desc).Distinct().ToList();
            else SearchProduto = DENUO.Sb1010s.Where(x => x.DELET != "*" && x.B1Msblql != "1" && x.B1Tipo != "KT" && x.B1Comerci=="C").Select(x => x.B1Desc).Distinct().ToList();
        }

        public IActionResult OnPostAsync(int Id,string Codigo,string NomePaciente,string NomeMedico,string NomeCliente,int Status,
            DateTime DataCirurgia,string Hora,string Procedimento,string Obs, IFormCollection Anexos01, IFormCollection Anexos02, 
            IFormCollection Anexos03, IFormCollection Anexos04, IFormCollection Anexos05, List<Produto> Produtos)
        {

            var dados = new DadosCirurgia
            {
                DataCriacao = DateTime.Now,
                Codigo = Codigo,
                NomePaciente = NomePaciente,
                NomeMedico = NomeMedico,
                NomeCliente = NomeCliente,
                Status = Status,
                DataCirurgia = DataCirurgia,
                ProcedimentosExec = Procedimento,
                ObsIntercorrencia = Obs,
                AgendamentoId = Id
            };

            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == Id);

            agendamento.StatusInstrumentador = 3;

            SGID.DadosCirurgias.Add(dados);
            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

            string Pasta = $"{_WEB.WebRootPath}/AnexosDadosCirurgia";

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

            Produtos.ForEach(produto =>
            {
                var ProdXAgenda = new DadosCirugiasProdutos
                {
                    DadosCirurgiaId = dados.Id, 
                    Quantidade = produto.Und,
                    Produto = produto.Item
                };

                SGID.DadosCirugiasProdutos.Add(ProdXAgenda);
                SGID.SaveChanges();

            });

            return LocalRedirect("/dashboard/0");
        }
    }
}
