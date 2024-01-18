using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.DTO;
using SGID.Models.Denuo;
using SGID.Data.Models;
using SGID.Models.Email;
using System.Net.Mail;
using SGID.Models.Inter;

namespace SGID.Pages
{
    [Authorize]
    public class DashBoardModel : PageModel
    {
        private readonly ILogger<DashBoardModel> _logger;
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        public List<Agendamentos> Agendamentos { get; set; }
        public List<RejeicaoMotivos> Rejeicoes { get; set; }

        private readonly IWebHostEnvironment _WEB;

        public int Total { get; set; }
        public int PendenteComercial { get; set; }
        public int RetornoCliente { get; set; }
        public int Respondidas { get; set; }
        public int NRespondidas { get; set; }
        public int Aprovadas { get; set; }
        public int Perdidas { get; set; }
        public int Emergencia { get; set; }

        public DashBoardModel(ILogger<DashBoardModel> logger,ApplicationDbContext sgid,TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter, IWebHostEnvironment wEB)
        {
            _logger = logger;
            SGID = sgid;
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
            _WEB = wEB;
        }
        public IActionResult OnGet(int id)
        {
           
            if (User.IsInRole("Instrumentador"))
            {
                return LocalRedirect("/instrumentador/dashboardinstrumentador");
            }
            else if (User.IsInRole("CoordenadorInstrumentador"))
            {
                return LocalRedirect("/dashboards/DashBoardGestorInstrumentador/0");
            }
            else if (User.IsInRole("Diretoria"))
            {
                return LocalRedirect("/dashboards/DashBoardMetas");
            }
            else if (User.IsInRole("Qualidade"))
            {
                return LocalRedirect("/dashboards/DashBoardQualidade");
            }
            else if (User.IsInRole("Financeiro") || User.IsInRole("Diretoria"))
            {
                return LocalRedirect("/relatorios/vexpenses/relatoriodespesas");
            }
            else if (User.IsInRole("SubDistribuidor"))
            {
                return LocalRedirect("/dashboards/DashBoardSubDistribuidor");
            }
            else if (User.IsInRole("Admin") || User.IsInRole("GestorVenda") || User.IsInRole("Venda") || User.IsInRole("Diretoria"))
            {

                return LocalRedirect($"/dashboards/DashBoard/{id}");
            }
            else if (User.IsInRole("GestorComercial"))
            {
                if (User.Identity.Name.Split("@")[1].ToUpper()=="INTERMEDIC.COM.BR")
                {
                    return LocalRedirect("/dashboards/DashBoardGestorComercial/01");
                }
                else
                {
                    return LocalRedirect("/dashboards/DashBoardGestorComercial/03");
                }       
                
            }
            else if (User.IsInRole("Comercial"))
            {
                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {
                    return LocalRedirect("/dashboards/DashBoardComercial/01");
                }
                else
                {
                    return LocalRedirect("/dashboards/DashBoardComercial/03");
                }

            }
            else if (User.IsInRole("RH"))
            {
                return LocalRedirect("/rh/listargestaopessoal");
            }
            else if (User.IsInRole("Patrimonio"))
            {
                return LocalRedirect("/formularios/patrimonio/relatoriopatrimonio/01");
            }
            else
            {
                return LocalRedirect("/relatorios/controladoria/conserto");
            }
        }
    }
}
