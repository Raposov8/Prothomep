using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.DTO;

namespace SGID.Pages.Instrumentador
{
    [Authorize(Roles = "Admin,Instrumentador,Diretoria")]
    public class DashBoardInstrumentadorModel : PageModel
    {
        public ApplicationDbContext SGID { get; set; }

        public int MeusAgendamentos { get; set; }

        public int Recusados { get; set; }

        public DashBoardInstrumentadorModel(ApplicationDbContext sGID)
        {
            SGID = sGID;
        }

        public void OnGet()
        {
            if (User.IsInRole("Admin") || User.IsInRole("Diretoria"))
            {
                MeusAgendamentos = SGID.Agendamentos.Count();
            }
            else
            {

                MeusAgendamentos = SGID.Agendamentos.Count();
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
                    x.Cliente,
                    x.Id,
                    x.Observacao,
                    x.Paciente
                }).FirstOrDefault(x => x.Id == id);

                var agenda = new
                {
                    Agendamento = agendamento,
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
