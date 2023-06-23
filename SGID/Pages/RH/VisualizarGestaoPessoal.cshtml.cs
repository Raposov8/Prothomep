using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data.ViewModel;
using SGID.Data;
using System.Net.Mail;
using Microsoft.AspNetCore.Authorization;

namespace SGID.Pages.RH
{
    [Authorize (Roles = "Admin,RH")]
    public class VisualizarGestaoPessoalModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        public SolicitacaoAcesso Acesso { get; set; }

        public VisualizarGestaoPessoalModel(ApplicationDbContext sgid) => SGID = sgid;

        public void OnGet(int id)
        {
            Acesso = SGID.SolicitacaoAcessos.FirstOrDefault(x => x.Id == id);
        }
    }
}
