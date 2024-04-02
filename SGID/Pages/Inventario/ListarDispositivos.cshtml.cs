using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;

namespace SGID.Pages.Inventario
{
    public class ListarDispositivosModel : PageModel
    {
        public ApplicationDbContext SGID { get; set; }

        public List<Dispositivo> Dispositivos { get; set; } = new List<Dispositivo>();

        public ListarDispositivosModel(ApplicationDbContext sgid)
        {
            SGID = sgid;
        }
        public void OnGet()
        {
            Dispositivos = SGID.Dispositivos.OrderBy(x=> x.TipoDispositivo).ToList();
        }

        public IActionResult OnPost(string Tipo)
        {
            if (Tipo != "") 
            {
                Dispositivos = SGID.Dispositivos.Where(x => x.TipoDispositivo == Tipo).OrderBy(x => x.TipoDispositivo).ToList();
            }
            else
            {
                Dispositivos = SGID.Dispositivos.OrderBy(x => x.TipoDispositivo).ToList();
            }
            return Page();
        }
    }
}
