using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;

namespace SGID.Pages.Estoque
{
    [Authorize]
	public class FormularioAvulsoPdfModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }

        public FormularioAvulso Formulario { get; set; }

        public List<FormularioAvulsoXProdutos> Produtos { get; set; }

        public FormularioAvulsoPdfModel(ApplicationDbContext sgid, TOTVSINTERContext inter, TOTVSDENUOContext denuo)
        {
            SGID = sgid;
            ProtheusInter = inter;
            ProtheusDenuo = denuo;
        }
        public void OnGet(int id)
        {
            try
            {
                Formulario = SGID.FormularioAvulsos.FirstOrDefault(x => x.Id == id);

                Produtos = SGID.FormularioAvulsoXProdutos.Where(x => x.FormularioId == Formulario.Id).ToList();
            }
            catch(Exception ex)
            {
                Logger.Log(ex, SGID, "FormularioAvulsoPdf", User.Identity.Name.Split("@")[0]);
            }
        }
    }
}
