using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;

namespace SGID.Pages.Estoque
{
	[Authorize]
	public class ListarFormulariosAvulsosModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        public List<FormularioAvulso> Formularios { get; set; }

        public ListarFormulariosAvulsosModel(ApplicationDbContext sgid) => SGID = sgid;
        public void OnGet()
        {
            Formularios = SGID.FormularioAvulsos.OrderByDescending(x => x.DataCriacao).Take(200).ToList();
        }

        public IActionResult OnPost(DateTime? Cirurgia,string NumAgendamento)
        {
            if(NumAgendamento != null)
            {
                Formularios = SGID.FormularioAvulsos.Where(x=> x.NumAgendamento == NumAgendamento).OrderByDescending(x => x.DataCriacao).Take(200).ToList();
            }
            else if(Cirurgia != null)
            {
                var DataFinal = Cirurgia.Value.AddDays(1);
                Formularios = SGID.FormularioAvulsos.Where(x => x.DataCirurgia >= Cirurgia && x.DataCirurgia < DataFinal).OrderByDescending(x => x.DataCriacao).ToList();
            }
            else
            {
                Formularios = SGID.FormularioAvulsos.OrderByDescending(x => x.DataCriacao).Take(200).ToList();
            }

            return Page();
        }
    }
}
