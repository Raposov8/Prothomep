using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Cirurgias;

namespace SGID.Pages.Formularios.Estoque
{
    public class OcorrenciaInterModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public List<string> Vendedores { get; set; }

        public NovoAgendamento Novo { get; set; } = new NovoAgendamento();

        public OcorrenciaInterModel(ApplicationDbContext sgid, TOTVSINTERContext inter)
        {
            SGID = sgid;
            ProtheusInter = inter;
        }

        public void OnGet()
        {
            Novo = new NovoAgendamento
            {
                Clientes = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Msblql != "1" && (x.A1Clinter == "C" || x.A1Clinter == "H" || x.A1Clinter == "M")).OrderBy(x => x.A1Nome).Select(x => x.A1Nreduz).ToList(),
                Medico = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Crm != "").OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                Hospital = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "H" && x.A1Msblql != "1").OrderBy(x => x.A1Nome).Select(x => x.A1Nreduz).ToList(),
            };

            Vendedores = ProtheusInter.Sa3010s.Where(x => x.DELET != "*" && x.A3Msblql != "1").Select(x => x.A3Nome).ToList();
        }

        public IActionResult OnPost(string Cliente,string Hospital, string Medico, string Paciente, string Agendamento, DateTime? Cirurgia,
            string Produto, string Descricao, string Ocorrencia, string Acao, string Procedente, string Cobrado, string Reposto,
            string Patrimonio, string DescPatri, string Vendedor, string Obs, DateTime? DataOcorrencia, int Quantidade)
        {
            var ocorrencia = new Ocorrencia
            {
                Empresa = "01",
                DataCriacao = DateTime.Now,
                DataOcorrencia = DataOcorrencia,
                Cliente = Cliente,
                Hospital = Hospital,
                Medico = Medico,
                Paciente = Paciente,
                Patrimonio = Patrimonio,
                DescPatri = DescPatri,
                Agendamento = Agendamento,
                Cirurgia = Cirurgia,
                Produto = Produto,
                Descricao = Descricao,
                Quantidade = Quantidade,
                Problema = Ocorrencia,
                Acao = Acao,
                Procedente = Procedente,
                Cobrado = Cobrado,
                Reposto = Reposto,
                Vendedor = Vendedor,
                Obs = Obs,
                UsuarioCriacao = User.Identity.Name.Split("@")[0].ToUpper()
            };


            SGID.Ocorrencias.Add(ocorrencia);
            SGID.SaveChanges();

            return LocalRedirect("/Formularios/Estoque/ListarOcorrencias");
        }

        public JsonResult OnGetPatrimonio(string Codigo)
        {
            var Descricao = ProtheusInter.Pa1010s.FirstOrDefault(x => x.Pa1Codigo == Codigo)?.Pa1Despat;

            return new JsonResult(Descricao);
        }
        public JsonResult OnGetProduto(string Codigo)
        {
            var Descricao = ProtheusInter.Sb1010s.FirstOrDefault(x => x.B1Cod == Codigo)?.B1Desc;

            return new JsonResult(Descricao);
        }
    }
}
