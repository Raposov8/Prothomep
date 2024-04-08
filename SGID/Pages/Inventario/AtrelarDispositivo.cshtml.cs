using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;

namespace SGID.Pages.Inventario
{
    public class AtrelarDispositivoModel : PageModel
    {
        public ApplicationDbContext SGID { get; set; }

        public AtrelarDispositivoModel(ApplicationDbContext sgid)
        {
            SGID = sgid;
        }
        public void OnGet(int Id)
        {
        }
        public IActionResult OnPost(int Id,string Nome)
        {
            var UsuarioDisp = new UsuarioDispositivo
            {
                Ativo = true,
                DispositivoId = Id,
                NomeUsuario = Nome
            };

            var OldUsers = SGID.UsuarioDispositivos.Where(x => x.DispositivoId == Id).ToList();

            OldUsers.ForEach(x =>
            {
                x.Ativo = false;
                SGID.UsuarioDispositivos.Update(x);
            });

            SGID.UsuarioDispositivos.Add(UsuarioDisp);
            SGID.SaveChanges();

            return LocalRedirect("/inventario/listardispositivos");
        }
    }
}
