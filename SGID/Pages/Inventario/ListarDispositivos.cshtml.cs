using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models;

namespace SGID.Pages.Inventario
{
    public class ListarDispositivosModel : PageModel
    {
        public ApplicationDbContext SGID { get; set; }

        public List<DispositivosAtivos> Dispositivos { get; set; } = new List<DispositivosAtivos>();

        public ListarDispositivosModel(ApplicationDbContext sgid)
        {
            SGID = sgid;
        }
        public void OnGet()
        {
            Dispositivos = (from Disp in SGID.Dispositivos
                            join User in SGID.UsuarioDispositivos on Disp.Id equals User.DispositivoId into st
                            from a in st.DefaultIfEmpty()
                            where a.Ativo != false
                            select new DispositivosAtivos
                            {
                                Id = Disp.Id,
                                TipoDispositivo = Disp.TipoDispositivo,
                                Imei = Disp.Imei,
                                Modelo = Disp.Modelo,
                                Nome = Disp.Nome,
                                NomeUsuario = a.NomeUsuario
                            }).ToList();
        }

        public IActionResult OnPost(string Tipo)
        {
            if (Tipo != "") 
            {
                Dispositivos = (from Disp in SGID.Dispositivos
                                join User in SGID.UsuarioDispositivos on Disp.Id equals User.DispositivoId into st
                                from a in st.DefaultIfEmpty()
                                where a.Ativo != false && Disp.TipoDispositivo == Tipo
                                select new DispositivosAtivos
                                {
                                    Id = Disp.Id,
                                    TipoDispositivo = Disp.TipoDispositivo,
                                    Imei = Disp.Imei,
                                    Modelo = Disp.Modelo,
                                    Nome = Disp.Nome,
                                    NomeUsuario = a.NomeUsuario
                                }).ToList();
            }
            else
            {
                Dispositivos = (from Disp in SGID.Dispositivos
                                join User in SGID.UsuarioDispositivos on Disp.Id equals User.DispositivoId into st
                                from a in st.DefaultIfEmpty()
                                where a.Ativo != false
                                select new DispositivosAtivos
                                {
                                    Id = Disp.Id,
                                    TipoDispositivo = Disp.TipoDispositivo,
                                    Imei = Disp.Imei,
                                    Modelo = Disp.Modelo,
                                    Nome = Disp.Nome,
                                    NomeUsuario = a.NomeUsuario
                                }).ToList();
            }
            return Page();
        }
    }
}
