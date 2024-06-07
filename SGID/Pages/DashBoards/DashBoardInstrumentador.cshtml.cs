using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.DashBoards
{
    [Authorize]
    public class DashBoardInstrumentadorModel : PageModel
    {
        public ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext DENUO { get; set; }
        private TOTVSINTERContext INTER { get; set; }

        public int MeusAgendamentos { get; set; }

        public int Recusados { get; set; }

        public DashBoardInstrumentadorModel(ApplicationDbContext sGID,TOTVSDENUOContext denuo,TOTVSINTERContext inter)
        {
            SGID = sGID;
            INTER = inter;
            DENUO = denuo;
        }

        public void OnGet()
        {
            if (User.IsInRole("Admin") || User.IsInRole("Diretoria"))
            {
                MeusAgendamentos = SGID.Agendamentos.Where(x=> x.StatusPedido == 7).Count();
            }
            else
            {

                MeusAgendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 7).Count();
            }
        }

        public JsonResult OnGetEvents()
        {
            try
            {
                if (User.IsInRole("Admin") || User.IsInRole("Diretoria"))
                {

                    var agendamentos = SGID.Agendamentos.ToList();

                    var events = new List<EventViewModel>();

                    agendamentos.ForEach(x =>
                    {
                        events.Add(new EventViewModel()
                        {
                            Id = x.Id,
                            Title = x.Paciente,
                            Start = $"{x.DataCirurgia:MM/dd/yyyy}",
                            AllDay = false,
                        });
                        //color = "yellow"
                    });

                    return new JsonResult(events.ToArray());
                }
                else
                {

                    var agendamentos = SGID.Agendamentos.ToList();

                    var events = new List<EventViewModel>();

                    agendamentos.ForEach(x =>
                    {
                        events.Add(new EventViewModel()
                        {
                            Id = x.Id,
                            Title = x.Paciente,
                            Start = $"{x.DataCirurgia:MM/dd/yyyy}",
                            AllDay = false,
                        });
                        //color = "yellow"
                    });

                    return new JsonResult(events.ToArray());
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "DashBoardInstrumentador, Events", user);

            }

            return new JsonResult("");
        }

        public JsonResult OnGetDetails(int id)
        {
            try
            {
                var agendamento = SGID.Agendamentos.Select(x => new
                {
                    x.Hospital,
                    x.Id,
                    x.Observacao,
                    x.Paciente,
                    x.Empresa,
                    x.CodHospital
                }).FirstOrDefault(x => x.Id == id);


                var endereco = "";

                if (agendamento.Empresa == "01") endereco = INTER.Sa1010s.FirstOrDefault(x => x.A1Cod == agendamento.CodHospital).A1End;
                else endereco = DENUO.Sa1010s.FirstOrDefault(x => x.A1Cod == agendamento.CodHospital).A1End;


                var agenda = new
                {
                    Agendamento = agendamento,
                    endereco
                };

                return new JsonResult(agenda);
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "DashBoardInstrumentador Details", user);
            }

            return new JsonResult("");
        }

        public class EventViewModel
        {
            public int Id { get; set; }
            public string Title { get; set; }
            public string Start { get; set; }
            public string End { get; set; }
            public bool AllDay { get; set; }
            //public string color { get; set; }
        }
    }
}
