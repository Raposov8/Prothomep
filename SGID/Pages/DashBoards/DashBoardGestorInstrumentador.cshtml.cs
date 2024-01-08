using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;

namespace SGID.Pages.DashBoards
{
    public class DashBoardGestorInstrumentadorModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        public List<Agendamentos> Agendamentos { get; set; } = new List<Agendamentos>();

        public int Agendadas { get; set; }
        public int Cirurgias { get; set; }
        public int Andamentos { get; set; }
        public int Concluidas { get; set; }
        public int Canceladas { get; set; }

        public DashBoardGestorInstrumentadorModel(ApplicationDbContext sgid)
        {
            SGID = sgid;

        }
        public void OnGet(int id)
        {

            else if (id == 2)
            {
                Agendamentos = ;
            }
            else if (id == 3)
            {
                Agendamentos = ;
            }
            else
            {
                Agendamentos = 
            }

            Agendamentos = id switch
            {
                0 => SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusInstrumentador == 0).ToList(),
                //Pendente Comercial Ida e volta da cotação
                1 => SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusInstrumentador == 1).ToList(),
                //Retorno Cliente
                2 => SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusInstrumentador == 2).ToList(),
                //NRespondidas
                3 => SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusInstrumentador == 3).ToList(),
                //
                _ => SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusInstrumentador == 4).ToList(),
            };


            Cirurgias = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusInstrumentador == 0).Count();

            Agendadas = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusInstrumentador == 1).Count();

            Andamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusInstrumentador == 2).Count();

            Concluidas = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusInstrumentador == 3).Count();

            Canceladas = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusInstrumentador == 4).Count();
        }

        public JsonResult OnGetInstrumentador(string Empresa)
        {

            if (Empresa == "01")
            {
                var Instrumentadores = SGID.Instrumentadores.Where(x => x.EmpresaProtheus == "INTERMEDIC" || x.EmpresaProtheus == null).Select(x => new { x.Id,x.Nome }).OrderBy(x=> x.Nome).ToList();

                return new JsonResult(Instrumentadores);
            }
            else
            {
                var Instrumentadores = SGID.Instrumentadores.Where(x => x.EmpresaProtheus == "DENUO" || x.EmpresaProtheus == null).Select(x => new { x.Id, x.Nome }).OrderBy(x => x.Nome).ToList();

                return new JsonResult(Instrumentadores);
            }

            
        }

        public JsonResult OnPostEnviar(int IdAgenda,int InstruAgenda, string Observacao)
        {

            var Instrumentador = SGID.Instrumentadores.FirstOrDefault(x => x.Id == InstruAgenda);

            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == IdAgenda);

            agendamento.Instrumentador = Instrumentador.Nome;
            agendamento.UsuarioGestorInstrumentador = User.Identity.Name.Split("@")[0].ToUpper();
            agendamento.StatusInstrumentador = 1;

            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();


            return new JsonResult("");
        }

        public JsonResult OnPostComercar(int IdAgenda, string Longitude ,string Latitude)
        {

            var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == IdAgenda);

            agendamento.Longitude = Longitude;
            agendamento.Latitude = Latitude;
            agendamento.StatusInstrumentador = 2;

            SGID.Agendamentos.Update(agendamento);
            SGID.SaveChanges();


            return new JsonResult("");
        }
    }
}
