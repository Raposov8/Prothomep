using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Cirurgias;
using SGID.Models.Denuo;

namespace SGID.Pages.Formularios.Estoque
{
    public class OcorrenciaModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }

        public List<string> Vendedores { get; set; }

        public NovoAgendamento Novo { get; set; } = new NovoAgendamento();

        public OcorrenciaModel(ApplicationDbContext sgid,TOTVSDENUOContext denuo)
        {
            SGID = sgid;
            ProtheusDenuo = denuo;
        }

        public void OnGet()
        {
            Novo = new NovoAgendamento
            {
                Clientes = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Msblql != "1" && (x.A1Clinter == "C" || x.A1Clinter == "H" || x.A1Clinter == "M")).OrderBy(x => x.A1Nome).Select(x => x.A1Nreduz).ToList(),
                Medico = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Crm != "").OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList(),
                Hospital = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "H" && x.A1Msblql != "1").OrderBy(x => x.A1Nome).Select(x => x.A1Nreduz).ToList(),
            };

            Vendedores = ProtheusDenuo.Sa3010s.Where(x => x.DELET != "*" && x.A3Msblql != "1").Select(x => x.A3Nome).ToList();
        }

        public IActionResult OnPost(string Cliente,string Hospital,string Medico,string Paciente,string Agendamento,DateTime? Cirurgia,
            string Produto,string Descricao,string Ocorrencia, string Acao,string Procedente,string Cobrado,string Reposto,
            string Patrimonio, string DescPatri,string Vendedor,string Obs,DateTime? DataOcorrencia,int Quantidade,string Armazem)
        {
            var ocorrencia = new Ocorrencia
            {
                Empresa = "03",
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
                Armazem = Armazem,
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
            var Descricao = (from PA10 in ProtheusDenuo.Pa1010s
                             join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                             from c in sr.DefaultIfEmpty()
                             join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                             from a in st.DefaultIfEmpty()
                             where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                             && c.DELET != "*" && a.DELET != "*"
                             && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                             select new
                             {
                                 PA10.Pa1Codigo,
                                 PA10.Pa1Despat
                             }).Distinct().FirstOrDefault(x => x.Pa1Codigo == Codigo)?.Pa1Despat;

            return new JsonResult(Descricao);
        }
        public JsonResult OnGetProduto(string Codigo)
        {
            var Descricao = ProtheusDenuo.Sb1010s.FirstOrDefault(x => x.B1Cod == Codigo)?.B1Desc;

            return new JsonResult(Descricao);
        }
    }
}
