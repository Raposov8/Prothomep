using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Inter;

namespace SGID.Pages.Formularios.Estoque
{
    public class OcorrenciaInterModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public List<string> Vendedores { get; set; }

        public OcorrenciaInterModel(ApplicationDbContext sgid, TOTVSINTERContext inter)
        {
            SGID = sgid;
            ProtheusInter = inter;
        }

        public void OnGet()
        {
            Vendedores = ProtheusInter.Sa3010s.Where(x => x.DELET != "*" && x.A3Msblql != "1").Select(x => x.A3Nome).ToList();
        }

        public IActionResult OnPost(string Cliente, string Medico, string Paciente, string Agendamento, DateTime? Cirurgia,
            string Produto, string Descricao, string Ocorrencia, string Acao, string Procedente, string Cobrado, string Reposto,
            string Patrimonio, string DescPatri, string Vendedor, string Obs, DateTime? DataOcorrencia, int Quantidade)
        {
            var ocorrencia = new Ocorrencia
            {
                Empresa = "01",
                DataCriacao = DateTime.Now,
                DataOcorrencia = DataOcorrencia,
                Cliente = Cliente,
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
                Obs = Obs
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
