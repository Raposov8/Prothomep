using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;

namespace SGID.Pages.Cotacoes
{
    [Authorize(Roles = "Admin,GestorComercial,Comercial")]
    public class DashBoardCotacoesModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        public List<Agendamentos> Agendamentos { get; set; }
        public List<RejeicaoMotivos> Rejeicoes { get; set; }

        public int Respondidas { get; set; }
        public int NRespondidas { get; set; }

        public DashBoardCotacoesModel(ApplicationDbContext sgid)
        {
            SGID = sgid;
        }

        public void OnGet(int id)
        {
            if (User.IsInRole("Admin"))
            {
                if (id == 0)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).ToList();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).Count();
                }
                else if (id == 1)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1).ToList();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).Count();
                }
                else if (id == 2)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 2).ToList();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).Count();
                }
            }
            else if (User.IsInRole("GestorComercial"))
            {
                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {
                    if (id == 0)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").ToList();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.Empresa == "01").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                    }
                    else if (id == 1)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.Empresa == "01").ToList();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.Empresa == "01").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                    }
                    else if (id == 2)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 2 && x.Empresa == "01").ToList();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.Empresa == "01").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                    }
                }
                else
                {
                    if (id == 0)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").ToList();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.Empresa == "03").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                    }
                    else if (id == 1)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.Empresa == "03").ToList();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.Empresa == "03").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                    }
                    else if (id == 2)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 2 && x.Empresa == "03").ToList();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.Empresa == "03").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                    }
                }

            }
            else
            {
                var user = User.Identity.Name.Split("@")[0].ToUpper();
                if (id == 0)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.VendedorLogin == user).ToList();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.VendedorLogin == user).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.VendedorLogin == user).Count();
                }
                else if (id == 1)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.VendedorLogin == user).ToList();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.VendedorLogin == user).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.VendedorLogin == user).Count();
                }
                else if (id == 2)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 2 && x.VendedorLogin == user).ToList();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.VendedorLogin == user).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.VendedorLogin == user).Count();
                }
            }

            Rejeicoes = SGID.RejeicaoMotivos.ToList();
        }
    }
}
