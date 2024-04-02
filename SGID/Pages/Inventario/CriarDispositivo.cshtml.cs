using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;

namespace SGID.Pages.Inventario
{
    public class CriarDispositivoModel : PageModel
    {
        public ApplicationDbContext SGID { get; set; }

        public CriarDispositivoModel(ApplicationDbContext sGID)
        {
            SGID = sGID;
        }

        public void OnGet()
        {

        }

        public IActionResult OnPost(string Nome,string Modelo,string Imei,string TipoDispositivo,double Valor)
        {

            var Dispositivo = new Dispositivo
            {
                DataCadastro = DateTime.Now,
                Nome = Nome,
                Imei = Imei,
                Modelo = Modelo,
                TipoDispositivo = TipoDispositivo,
                Valor = Valor,
                UsuarioCriacao = User.Identity.Name.Split('@')[0]
            };

            SGID.Dispositivos.Add(Dispositivo);
            SGID.SaveChanges();

            return LocalRedirect("/inventario/listardispositivos");
        }
    }
}
