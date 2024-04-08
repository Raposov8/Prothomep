using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;

namespace SGID.Pages.Inventario
{
    public class EditarDispositivoModel : PageModel
    {
        public ApplicationDbContext SGID { get; set; }

        public Dispositivo Editar { get; set; } = new Dispositivo();

        public EditarDispositivoModel(ApplicationDbContext sGID)
        {
            SGID = sGID;
        }

        public void OnGet(int Id)
        {
            Editar = SGID.Dispositivos.First(x => x.Id == Id);
        }

        public IActionResult OnPost(int Id,string Nome, string Modelo, string Imei, string TipoDispositivo, double Valor)
        {

            Editar = SGID.Dispositivos.First(x => x.Id == Id);

            Editar.Nome = Nome;
            Editar.Imei = Imei;
            Editar.Modelo = Modelo;
            Editar.TipoDispositivo = TipoDispositivo;
            Editar.Valor = Valor;
            Editar.DataAlteracao = DateTime.Now;
            Editar.UsuarioAlteracao = User.Identity.Name.Split("@")[0].ToUpper();
           

            SGID.Dispositivos.Update(Editar);
            SGID.SaveChanges();

            return LocalRedirect("/inventario/listardispositivos");
        }
    }
}
