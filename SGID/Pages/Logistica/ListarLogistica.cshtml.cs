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
        public int EmRota { get; set; } = 0;
        public int Entregue { get; set; } = 0;
        public int Retorno { get; set; } = 0;

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
                    //EmRota
                    2 => SGID.Agendamentos.Where(x => x.StatusLogistica == 1 && x.Empresa == "01")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Entregue
                    3 => SGID.Agendamentos.Where(x => (x.StatusLogistica == 2 || x.StatusLogistica == 3) && x.Empresa == "01")
                         .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Retorno
                    4 => SGID.Agendamentos.Where(x => x.StatusLogistica == 4 && x.Empresa == "01")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    _ => new List<Agendamentos>()
                };

                Aprovadas = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.DataCirurgia != null && x.StatusLogistica == 0 && x.Empresa == "01").Count();

                EmRota = SGID.Agendamentos.Where(x => x.StatusLogistica == 1 && x.Empresa == "01").Count();

                Entregue = SGID.Agendamentos.Where(x => (x.StatusLogistica == 2 || x.StatusLogistica == 3) && x.Empresa == "01").Count();

                Retorno = SGID.Agendamentos.Where(x => x.StatusLogistica == 4 && x.Empresa == "01").Count();
            }
            else
            {
                Agendamentos = id switch
                {
                    //Aprovadas
                    1 => SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.DataCirurgia != null && x.StatusLogistica == 0 && x.Empresa =="03")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //EmRota
                    2 => SGID.Agendamentos.Where(x => x.StatusLogistica == 1 && x.Empresa == "03")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Entregue
                    3 => SGID.Agendamentos.Where(x => (x.StatusLogistica == 2 || x.StatusLogistica == 3) && x.Empresa == "03")
                         .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    //Retorno
                    4 => SGID.Agendamentos.Where(x => x.StatusLogistica == 4 && x.Empresa == "03")
                        .OrderBy(x => x.Tipo).ThenBy(x => x.DataCirurgia).ToList(),
                    _ => new List<Agendamentos>()
                };

                Aprovadas = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.DataCirurgia != null && x.StatusLogistica == 0 && x.Empresa == "03").Count();

                EmRota = SGID.Agendamentos.Where(x => x.StatusLogistica == 1 && x.Empresa == "03").Count();

                Entregue = SGID.Agendamentos.Where(x => (x.StatusLogistica == 2 || x.StatusLogistica == 3) && x.Empresa == "03").Count();

                Retorno = SGID.Agendamentos.Where(x => x.StatusLogistica == 4 && x.Empresa == "03").Count();
            }

            Empresa = empresa;

            Rejeicoes = SGID.RejeicaoMotivos.ToList();
        }

        public JsonResult OnGetEnviar(int IdA)
        {
            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == IdA);

            agendamento.UsuarioLogistica = User.Identity.Name.Split("@")[0].ToUpper();
            agendamento.StatusLogistica = 1;

            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();

            return new JsonResult("");
        }

    }
}
