using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Inter;

namespace SGID.Pages.Cirurgias
{
    [Authorize(Roles = "Admin,GestorVenda,Venda,Diretoria")]
    public class VisualizarRecursoModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext Protheus { get; set; }
        public Pah010 Instrumentador { get; set; }

        public VisualizarRecursoModel(ApplicationDbContext sgid, TOTVSINTERContext protheus)
        { 
            SGID = sgid;
            Protheus = protheus;
        }
        public void OnGet(int Id)
        {
            Instrumentador = Protheus.Pah010s.FirstOrDefault(x => x.RECNO == Id);
        }
    }
}
