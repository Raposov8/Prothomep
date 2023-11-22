using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
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
        public int Canceladas { get; set; }

        public DashBoardCotacoesModel(ApplicationDbContext sgid)
        {
            SGID = sgid;
        }

        public void OnGet(int id)
        {
            if (User.IsInRole("Admin") || User.IsInRole("Diretoria"))
            {
                if (id == 0)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) ).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4).Count();
                }
                else if (id == 1)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  ).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4).Count();
                }
                else if (id == 2)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) ).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4).Count();
                }
                else if (id == 3)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.StatusLogistica == 1).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) ).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4).Count();
                }
                else if (id == 4)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 1).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) ).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4).Count();
                }
                else if(id == 5)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) ).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4).Count();
                }
                else
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1) ).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4).Count();
                }
            }
            else if (User.IsInRole("GestorComercial"))
            {
                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {

                    if (id == 0)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "01").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "01").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.StatusPedido!=7 && x.Empresa == "01").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "01").Count();
                    }
                    else if (id == 1)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.StatusPedido!=7 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "01").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "01").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.StatusPedido!=7 && x.Empresa == "01").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "01").Count();
                    }
                    else if (id == 2)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "01").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "01").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.StatusPedido!=7 && x.Empresa == "01").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "01").Count();
                    }
                    else if (id == 3)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.StatusLogistica == 1 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "01").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "01").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.StatusPedido!=7 && x.Empresa == "01").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "01").Count();
                    }
                    else if (id == 4)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "01").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "01").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.StatusPedido!=7 && x.Empresa == "01").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "01").Count();
                    }
                    else if (id == 5)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "01").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "01").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.StatusPedido!=7 && x.Empresa == "01").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "01").Count();
                    }
                    else if (id == 6)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "01").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "01").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.Empresa == "01").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "01").Count();
                    }
                    else
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "01").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "01").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "01").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "01").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.Empresa == "01").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "01").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "01").Count();
                    }
                }
                else
                {
                    if (id == 0)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "03").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "03").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.Empresa == "03").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "03").Count();
                    }
                    else if (id == 1)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "03").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "03").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.Empresa == "03").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "03").Count();
                    }
                    else if (id == 2)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "03").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "03").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.Empresa == "03").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "03").Count();
                    }
                    else if (id == 3)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.StatusLogistica == 1 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "03").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "03").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.Empresa == "03").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "03").Count();
                    }
                    else if (id == 4)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "03").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "03").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.Empresa == "03").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "03").Count();
                    }
                    else if (id == 5)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "03").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "03").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.Empresa == "03").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "03").Count();
                    }
                    else if (id == 6)
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "03").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "03").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.Empresa == "03").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "03").Count();
                    }
                    else
                    {
                        Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "03").OrderByDescending(x => x.Id).ToList();
                        Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.Empresa == "03").Count();
                        AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.Empresa == "03").Count();
                        RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 && x.Empresa == "03").Count();
                        Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 && x.Empresa == "03").Count();
                        NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 && x.Empresa == "03").Count();
                        Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.Empresa == "03").Count();
                    }
                }

            }
            else
            {
                var user = User.Identity.Name.Split("@")[0].ToUpper();

                if (id == 0)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 &&  x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 &&  x.VendedorLogin == user).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  &&  x.VendedorLogin == user).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 &&  x.VendedorLogin == user).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 &&  x.VendedorLogin == user).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 &&  x.VendedorLogin == user).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 &&  x.VendedorLogin == user).Count();
                }
                else if (id == 1)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 &&  x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 &&  x.VendedorLogin == user).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  &&  x.VendedorLogin == user).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 &&  x.VendedorLogin == user).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 &&  x.VendedorLogin == user).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 &&  x.VendedorLogin == user).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 &&  x.VendedorLogin == user).Count();
                }
                else if (id == 2)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 &&  x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 &&  x.VendedorLogin == user).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  &&  x.VendedorLogin == user).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 &&  x.VendedorLogin == user).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 &&  x.VendedorLogin == user).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 &&  x.VendedorLogin == user).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 &&  x.VendedorLogin == user).Count();
                }
                else if (id == 3)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 && x.StatusLogistica == 1 &&  x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 &&  x.VendedorLogin == user).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  &&  x.VendedorLogin == user).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 &&  x.VendedorLogin == user).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 &&  x.VendedorLogin == user).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 &&  x.VendedorLogin == user).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 &&  x.VendedorLogin == user).Count();
                }
                else if (id == 4)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  &&  x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 &&  x.VendedorLogin == user).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  &&  x.VendedorLogin == user).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 &&  x.VendedorLogin == user).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 &&  x.VendedorLogin == user).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 &&  x.VendedorLogin == user).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 &&  x.VendedorLogin == user).Count();
                }
                else if (id == 5)
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 &&  x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 &&  x.VendedorLogin == user).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  &&  x.VendedorLogin == user).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 7 &&  x.VendedorLogin == user).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7 &&  x.VendedorLogin == user).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2 &&  x.VendedorLogin == user).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 &&  x.VendedorLogin == user).Count();
                }
                else
                {
                    Agendamentos = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 &&  x.VendedorLogin == user).OrderByDescending(x => x.Id).ToList();
                    Aprovadas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 3 &&  x.VendedorLogin == user).Count();
                    AguardandoConfirmacao = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && (x.StatusLogistica == 0 || x.StatusLogistica == 1)  && x.VendedorLogin == user).Count();
                    RespondidasEstoque = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 7 && x.StatusLogistica == 1  && x.VendedorLogin == user).Count();
                    Respondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido != 3 && x.StatusPedido!=7  && x.VendedorLogin == user).Count();
                    NRespondidas = SGID.Agendamentos.Where(x => x.StatusCotacao == 0 && x.StatusPedido == 2  && x.VendedorLogin == user).Count();
                    Canceladas = SGID.Agendamentos.Where(x => x.StatusCotacao == 1 && x.StatusPedido == 4 && x.VendedorLogin == user).Count();
                }
            }

            Rejeicoes = SGID.RejeicaoMotivos.ToList();
        }
    }
}
