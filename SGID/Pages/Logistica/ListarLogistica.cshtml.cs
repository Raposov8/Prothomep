using DocumentFormat.OpenXml.Bibliography;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;

namespace SGID.Pages.Logistica
{
    [Authorize]
    public class ListarLogisticaModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public List<Agendamentos> Agendamentos { get; set; }
        public List<RejeicaoMotivos> Rejeicoes { get; set; }

        public string Empresa { get; set; } = "";
        public int Aprovadas { get; set; } = 0;

        public int Confirmadas { get; set; } = 0;
        public int Pendente { get; set; } = 0;
        public int Separacao { get; set; } = 0;
        public int EmRota { get; set; } = 0;
        public int Entregue { get; set; } = 0;
        public int Retorno { get; set; } = 0;
        public int Inspecao { get; set; } = 0;

        public ListarLogisticaModel(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter,ApplicationDbContext sgid)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
            SGID = sgid;
        }

        public void OnGet(string empresa,int id)
        {
            if(empresa == "01")
            {
                Agendamentos = id switch
                {
                    //Aprovadas
                    1 => SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.DataCirurgia != null && x.StatusLogistica == 0 && x.Empresa == "01")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Pendente
                    2 => SGID.Agendamentos.Where(x => (x.StatusLogistica == 1 || x.StatusLogistica == 7) && x.Empresa == "01")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Separação
                    3 => SGID.Agendamentos.Where(x => x.StatusLogistica == 2 && x.Empresa == "01")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //EmRota
                    4 => SGID.Agendamentos.Where(x => x.StatusLogistica == 3 && x.Empresa == "01")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Entregue
                    5 => SGID.Agendamentos.Where(x => x.StatusLogistica == 4 && x.Empresa == "01")
                         .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Retorno
                    6 => SGID.Agendamentos.Where(x => x.StatusLogistica == 5 && x.Empresa == "01")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Inspeção
                    7 => SGID.Agendamentos.Where(x => x.StatusLogistica == 6 && x.Empresa == "01")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Confirmadas
                    8 => SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.DataCirurgia != null && x.StatusLogistica == 0 && x.Empresa == "01")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    _ => new List<Agendamentos>()
                };

                Aprovadas = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.DataCirurgia != null && x.StatusLogistica == 0 && x.Empresa == "01").Count();

                Confirmadas = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.DataCirurgia != null && x.StatusLogistica == 0 && x.Empresa == "01").Count();

                Pendente = SGID.Agendamentos.Where(x => (x.StatusLogistica == 1 || x.StatusLogistica == 7) && x.Empresa == "01").Count();

                Separacao = SGID.Agendamentos.Where(x => x.StatusLogistica == 2 && x.Empresa == "01").Count();

                EmRota = SGID.Agendamentos.Where(x => x.StatusLogistica == 3 && x.Empresa == "01").Count();

                Entregue = SGID.Agendamentos.Where(x => x.StatusLogistica == 4 && x.Empresa == "01").Count();

                Retorno = SGID.Agendamentos.Where(x => x.StatusLogistica == 5 && x.Empresa == "01").Count();

                Inspecao = SGID.Agendamentos.Where(x => x.StatusLogistica == 6 && x.Empresa == "01").Count();
            }
            else
            {
                Agendamentos = id switch
                {
                    //Aprovadas
                    1 => SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.DataCirurgia != null && x.StatusLogistica == 0 && x.Empresa == "03")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Pendente
                    2 => SGID.Agendamentos.Where(x => (x.StatusLogistica == 1 || x.StatusLogistica == 7)&& x.Empresa == "03")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Separação
                    3 => SGID.Agendamentos.Where(x => x.StatusLogistica == 2 && x.Empresa == "03")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //EmRota
                    4 => SGID.Agendamentos.Where(x => x.StatusLogistica == 3 && x.Empresa == "03")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Entregue
                    5 => SGID.Agendamentos.Where(x => x.StatusLogistica == 4 && x.Empresa == "03")
                         .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Retorno
                    6 => SGID.Agendamentos.Where(x => x.StatusLogistica == 5 && x.Empresa == "03")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Inspeção
                    7 => SGID.Agendamentos.Where(x => x.StatusLogistica == 6 && x.Empresa == "03")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Confirmadas
                    8 => SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.DataCirurgia != null && x.StatusLogistica == 0 && x.Empresa == "03")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    _ => new List<Agendamentos>()
                };

                Aprovadas = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.DataCirurgia != null && x.StatusLogistica == 0 && x.Empresa == "03").Count();

                Confirmadas = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.DataCirurgia != null && x.StatusLogistica == 0 && x.Empresa == "03").Count();

                Pendente = SGID.Agendamentos.Where(x => (x.StatusLogistica == 1 || x.StatusLogistica == 7) && x.Empresa == "03").Count();

                Separacao = SGID.Agendamentos.Where(x => x.StatusLogistica == 2 && x.Empresa == "03").Count();

                EmRota = SGID.Agendamentos.Where(x => x.StatusLogistica == 3 && x.Empresa == "03").Count();

                Entregue = SGID.Agendamentos.Where(x => x.StatusLogistica == 4 && x.Empresa == "03").Count();

                Retorno = SGID.Agendamentos.Where(x => x.StatusLogistica == 5 && x.Empresa == "03").Count();

                Inspecao = SGID.Agendamentos.Where(x => x.StatusLogistica == 6 && x.Empresa == "03").Count();
            }

            Empresa = empresa;

            Rejeicoes = SGID.RejeicaoMotivos.ToList();
        }

        public IActionResult OnGetEnviarComercial(int IdA)
        {
            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == IdA);

            //agendamento.UsuarioLogistica = User.Identity.Name.Split("@")[0].ToUpper();
            agendamento.StatusLogistica = 1;

            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

            return LocalRedirect($"/Logistica/ListarLogistica/{agendamento.Empresa}/2");
        }

        public IActionResult OnGetEnviarParaRota(int IdA)
        {
            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == IdA);

            //agendamento.UsuarioLogistica = User.Identity.Name.Split("@")[0].ToUpper();
            agendamento.StatusLogistica = 4;

            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

            return LocalRedirect($"/Logistica/ListarLogistica/${agendamento.Empresa}/5");
        }

        public IActionResult OnGetEnviarParaEntregue(int IdA)
        {
            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == IdA);

            //agendamento.UsuarioLogistica = User.Identity.Name.Split("@")[0].ToUpper();
            agendamento.StatusLogistica = 5;

            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

            return LocalRedirect($"/Logistica/ListarLogistica/${agendamento.Empresa}/6");
        }

        public IActionResult OnGetEnviarParaInspecao(int IdA)
        {
            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == IdA);

            //agendamento.UsuarioLogistica = User.Identity.Name.Split("@")[0].ToUpper();
            agendamento.StatusLogistica = 6;

            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

            return LocalRedirect($"/Logistica/ListarLogistica/${agendamento.Empresa}/7");
        }

        public IActionResult OnGetEnviarParaArquivo(int IdA)
        {
            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == IdA);

            //agendamento.UsuarioLogistica = User.Identity.Name.Split("@")[0].ToUpper();
            agendamento.StatusLogistica = 7;

            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

            return LocalRedirect($"/Logistica/ListarLogistica/${agendamento.Empresa}/7");
        }

    }
}
