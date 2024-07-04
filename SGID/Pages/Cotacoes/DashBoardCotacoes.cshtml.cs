using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;

namespace SGID.Pages.Cotacoes
{
    [Authorize(Roles = "Admin,GestorComercial,Comercial,Diretoria")]
    public class DashBoardCotacoesModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        public List<Agendamentos> Agendamentos { get; set; }
        public List<RejeicaoMotivos> Rejeicoes { get; set; }

        public int Aprovadas { get; set; }
        public int AguardandoConfirmacao { get; set; }
        public int RespondidasEstoque { get; set; }
        public int Respondidas { get; set; }
        public int NRespondidas { get; set; }
        public int Logistica { get; set; }
        public int Canceladas { get; set; }
        public string Empresa { get; set; }
        public bool HorarioEmergencia { get; set; }

        public DashBoardCotacoesModel(ApplicationDbContext sgid)
        {
            SGID = sgid;
        }

        public void OnGet(int id)
        {
            if (User.IsInRole("Admin") || User.IsInRole("Diretoria"))
            {
                switch (id)
                {
                    case 0:
                        {
                            Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).OrderByDescending(x => x.Id).ToList();
                            break;
                        }
                    case 1:
                        {
                            Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido != 7).OrderByDescending(x => x.Id).ToList();
                            break;
                        }
                    case 2:
                        {
                            Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 3).OrderByDescending(x => x.Id).ToList();
                            break;
                        }
                    case 3:
                        {
                            Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.StatusLogistica == 1).OrderByDescending(x => x.Id).ToList();
                            break;
                        }
                    case 4:
                        {
                            Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusLogistica == 1).OrderByDescending(x => x.Id).ToList();
                            break;
                        }
                    case 5:
                        {
                            Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusLogistica == 7).OrderByDescending(x => x.Id).ToList();
                            break;
                        }
                    case 6:
                        {
                            Agendamentos = SGID.Agendamentos.Where(x=> x.StatusPedido == 7 && x.StatusLogistica <= 6 && x.StatusLogistica>=2).OrderByDescending(x => x.Id).ToList();
                            break;
                        }
                    default:
                        {
                            Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 4).OrderByDescending(x => x.Id).ToList();
                            break;
                        }
                }

                Aprovadas = SGID.Agendamentos.Where(x => x.StatusPedido == 3).Count();
                AguardandoConfirmacao = SGID.Agendamentos.Where(x =>  x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)).Count();
                RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusLogistica == 7).Count();
                Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido != 7).Count();
                NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).Count();
                Logistica = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusLogistica <= 6 && x.StatusLogistica >= 2).Count();
                Canceladas = SGID.Agendamentos.Where(x => x.StatusPedido == 4).Count();
            }
            else if (User.IsInRole("GestorComercial"))
            {
                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {
                    switch (id)
                    {
                        case 0:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 1:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido != 7 && x.StatusPedido != 7 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 2:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 3:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x =>  x.StatusPedido == 3 && x.StatusLogistica == 1 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 4:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x =>  x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 5:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x =>  x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 6:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusLogistica <= 6 && x.StatusLogistica >= 2 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        default:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x =>  x.StatusPedido == 4 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                    }

                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.Empresa == "01").Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) && x.Empresa == "01").Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x =>  x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido != 7 && x.StatusPedido != 7 && x.Empresa == "01").Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                    Logistica = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusLogistica <= 6 && x.StatusLogistica >= 2 && x.Empresa == "01").Count();
                    Canceladas = SGID.Agendamentos.Where(x =>  x.StatusPedido == 4 && x.Empresa == "01").Count();
                }
                else
                {
                    switch (id) 
                    {
                        case 0:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 1:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido != 7 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 2:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 3:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.StatusLogistica == 1 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 4:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x =>  x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 5:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x =>  x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        case 6:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusLogistica <= 6 && x.StatusLogistica >= 2 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                        default:
                            {
                                Agendamentos = SGID.Agendamentos.Where(x =>  x.StatusPedido == 4 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                                break;
                            }
                    }

                    Aprovadas = SGID.Agendamentos.Where(x =>  x.StatusPedido == 3 && x.Empresa == "03").Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x =>  x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) && x.Empresa == "03").Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x =>  x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido != 7 && x.Empresa == "03").Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                    Logistica = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusLogistica <= 6 && x.StatusLogistica >= 2 && x.Empresa == "03").Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusPedido == 4 && x.Empresa == "03").Count();
                }
            }
            else
            {
                var user = User.Identity.Name.Split("@")[0].ToUpper();

                switch (id)
                {
                    case 0:
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                        break;
                    }
                    case 1:
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido != 7 && x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                        break;
                    }
                    case 2:
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                        break;
                    }
                    case 3:
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.StatusLogistica == 1 && x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                        break;
                    }
                    case 4:
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) && x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                        break;
                    }
                    case 5:
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusLogistica == 7 && x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                        break;
                    }
                    case 6:
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusLogistica <= 6 && x.StatusLogistica >= 2 && x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                        break;
                    }
                    default:
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusPedido == 4 && x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                        break;
                    }
                }

                Aprovadas = SGID.Agendamentos.Where(x => x.StatusPedido == 3 && x.VendedorLogin == user).Count();
                AguardandoConfirmacao = SGID.Agendamentos.Where(x =>  x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) && x.VendedorLogin == user).Count();
                RespondidasEstoque = SGID.Agendamentos.Where(x =>  x.StatusPedido == 7 && x.StatusLogistica == 7 && x.VendedorLogin == user).Count();
                Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido != 7 && x.VendedorLogin == user).Count();
                NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.VendedorLogin == user).Count();
                Logistica = SGID.Agendamentos.Where(x => x.StatusPedido == 7 && x.StatusLogistica <= 6 && x.StatusLogistica >= 2 && x.VendedorLogin == user).Count();
                Canceladas = SGID.Agendamentos.Where(x =>  x.StatusPedido == 4 && x.VendedorLogin == user).Count();
            }

            var Data = DateTime.Now;

            if(Data.Hour < 7 || Data.Hour > 18)
            {
                HorarioEmergencia = true;
            }

            HorarioEmergencia = true;

            Rejeicoes = SGID.RejeicaoMotivos.ToList();
        }
    }
}
