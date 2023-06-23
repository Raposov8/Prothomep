using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.Agendamento
{
    [Authorize]
    public class VisitasModel : PageModel
    {
        public TOTVSDENUOContext ProtheusDenuo { get; set; }
        public TOTVSINTERContext ProtheusInter { get; set; }
        public ApplicationDbContext SGID { get; set; }
        public string BaseUrl { get; set; }

        public List<string> Medicos { get; set; } = new List<string>();
        public List<string> Locais { get; set; } = new List<string> { "HOSPITAL","CLINICA","CONSULTORIO","CASA","HOTEL","LOCAL PUBLICO" };

        public VisitasModel(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter, ApplicationDbContext sGID)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
            SGID = sGID;
        }

        public void OnGet()
        {
            BaseUrl = $"{Request.Scheme}://{Request.Host}";

            string usuario = User.Identity.Name.Split("@")[0].ToUpper();

            if (User.IsInRole("Admin"))
            {
                Medicos = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend)).OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList();
                var medicos = SGID.VisitaClientes.Where(x => x.Empresa == "INTERMEDIC").Select(x => x.Medico).ToList();

                Medicos.AddRange(medicos);
                Medicos = Medicos.OrderBy(x => x).ToList();
            }
            else if (User.IsInRole("GestorComercial"))
            {
                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {
                    Medicos = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend)).OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList();

                    var medicos = SGID.VisitaClientes.Where(x => x.Empresa == "INTERMEDIC").Select(x => x.Medico).ToList();

                    Medicos.AddRange(medicos);
                    Medicos = Medicos.OrderBy(x => x).ToList();
                }
                else
                {
                    Medicos = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend)).OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList();
                    var medicos = SGID.VisitaClientes.Where(x => x.Empresa == "DENUO").Select(x => x.Medico).ToList();

                    Medicos.AddRange(medicos);
                    Medicos = Medicos.OrderBy(x => x).ToList();
                }
            }
            else
            {
                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {
                    Medicos = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend)).OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList();

                    var medicos = SGID.VisitaClientes.Where(x => x.Empresa == "INTERMEDIC").Select(x => x.Medico).ToList();

                    Medicos.AddRange(medicos);
                    Medicos = Medicos.OrderBy(x => x).ToList();
                }
                else
                {
                    Medicos = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend)).OrderBy(x => x.A1Nome).Select(x => x.A1Nome).ToList();
                    var medicos = SGID.VisitaClientes.Where(x => x.Empresa == "DENUO").Select(x => x.Medico).ToList();

                    Medicos.AddRange(medicos);
                    Medicos = Medicos.OrderBy(x => x).ToList();
                }
            }
        }

        public JsonResult OnGetEvents()
        {
            try
            {

                string usuario = User.Identity.Name.Split("@")[0].ToUpper();

                List<Visitas> visitas = new List<Visitas>();
                if (User.IsInRole("Admin"))
                {
                    visitas = SGID.Visitas.ToList();
                }
                else if (User.IsInRole("GestorComercial"))
                {
                    if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                    {
                        var Equipe = SGID.Times.Where(x=> x.Lider == usuario.ToLower()).ToList();

                        Equipe.ForEach(x =>
                        {
                            var VisitaInte = SGID.Visitas.Where(c => c.Vendedor == x.Integrante.ToUpper()).ToList();

                            visitas.AddRange(VisitaInte);
                        });

                        var visita = SGID.Visitas.Where(c => c.Vendedor == usuario).ToList();

                        visitas.AddRange(visita);
                    }
                    else
                    {
                        var Equipe = SGID.Times.Where(x => x.Lider == usuario.ToLower()).ToList();

                        Equipe.ForEach(x =>
                        {
                            var VisitaInte = SGID.Visitas.Where(c => c.Vendedor == x.Integrante.ToUpper()).ToList();

                            visitas.AddRange(VisitaInte);
                        });

                        var visita = SGID.Visitas.Where(c => c.Vendedor == usuario).ToList();

                        visitas.AddRange(visita);
                    }
                }
                else
                {

                    visitas = SGID.Visitas.Where(x => x.Vendedor == usuario).ToList();
                }
               
                var events = new List<EventViewModel>();

                visitas.ForEach(x =>
                {
                    events.Add(new EventViewModel()
                    {
                        Id = x.Id,
                        Title = x.Medico,
                        Start = $"{x.DataHora:MM/dd/yyyy HH:mm}",
                        AllDay = false,
                    });
                    //color = "yellow"
                });
                return new JsonResult(events.ToArray());
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Visitas", user);
            }

            return new JsonResult("");
        }

        public JsonResult OnGetDetails(int id)
        {
            try
            {
                var visita = SGID.Visitas
                    .Select(x=> new
                    {
                        x.Id,
                        DataHora = x.DataHora.ToString("yyyy/MM/dd HH:mm"),
                        DataUltima = x.DataUltima != null ? x.DataUltima.Value.ToString("yyyy/MM/dd"):"Não há registro",
                        x.Local,
                        x.Assunto,
                        x.Observacao,
                        x.Endereco,
                        x.Medico,
                        x.Motvisita,
                        x.Bairro,
                        ResuVi = x.ResumoVisita,
                        Vendedor = x.Vendedor.Replace("."," "),

                    }).FirstOrDefault(x=> x.Id == id);

                return new JsonResult(visita);
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Visitas Details", user);
            }

            return new JsonResult("");
        }

        public IActionResult OnPost(string Medico, string Local, DateTime Data, DateTime DataUltima, string Assunto,
            string Endereco,string Bairro, string MotVis, string Obs,string ResuVisi)
        {
            try
            {
                var Visita = new Visitas
                {
                    DataCriacao = DateTime.Now,
                    Medico = Medico,
                    Local = Local,
                    DataHora = Data,
                    DataUltima = DataUltima,
                    Assunto = Assunto,
                    Endereco = Endereco,
                    Bairro = Bairro,
                    Motvisita = MotVis,
                    Observacao = Obs,
                    ResumoVisita = ResuVisi,
                    Vendedor = User.Identity.Name.Split("@")[0].ToUpper(),
                    Empresa = User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR" ? "INTERMEDIC":"DENUO"
                };

                SGID.Visitas.Add(Visita);
                SGID.SaveChanges();

                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {
                    if (ProtheusInter.Sa1010s.FirstOrDefault(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico) == null)
                    {

                        var VisitaCliente = new VisitaCliente
                        {
                            DataCriacao = DateTime.Now,
                            Medico = Medico,
                            Local = Local,
                            Assunto = Assunto,
                            Endereco = Endereco,
                            Bairro = Bairro,
                            Vendedor = User.Identity.Name.Split("@")[0].ToUpper(),
                            Empresa = "INTERMEDIC"
                        };


                        SGID.VisitaClientes.Add(VisitaCliente);
                        SGID.SaveChanges();
                    }
                }
                else
                {
                    if (ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico) == null)
                    {
                        var VisitaCliente = new VisitaCliente
                        {
                            DataCriacao = DateTime.Now,
                            Medico = Medico,
                            Local = Local,
                            Assunto = Assunto,
                            Endereco = Endereco,
                            Bairro = Bairro,
                            Vendedor = User.Identity.Name.Split("@")[0].ToUpper(),
                            Empresa = "DENUO"
                        };

                        SGID.VisitaClientes.Add(VisitaCliente);
                        SGID.SaveChanges();
                    }
                }

                return LocalRedirect("/agendamento/visitas");

            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovaVisita", user);
                return Page();
            }
        }

        public JsonResult OnPostResumo(int id, string Observacao, string Resumo)
        {
            try 
            {
                var visita = SGID.Visitas.FirstOrDefault(x => x.Id == id);

                visita.Observacao = Observacao;
                visita.ResumoVisita = Resumo;
                visita.DataAlteracao = DateTime.Now;

                SGID.Visitas.Update(visita);
                SGID.SaveChanges();

                return new JsonResult("");

            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovaVisita", user);
                return new JsonResult("");
            }
        }

        public JsonResult OnGetEndereco(string Medico)
        {

            if (User.IsInRole("Admin"))
            {
                var endereco = ProtheusInter.Sa1010s.FirstOrDefault(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico);

                var Quants = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Bairro == endereco.A1Bairro).Count();

                var Data = (SGID.Visitas.OrderByDescending(x => x.DataHora).FirstOrDefault(x => x.Medico == Medico && x.Empresa == "INTERMEDIC")?.DataHora) ?? DateTime.Now;
                var End = new
                {
                    Endereco = endereco.A1End,
                    Bairro = endereco.A1Bairro,
                    Quant = Quants,
                    Ultima = Data.ToString("yyyy-MM-dd")
                };

                return new JsonResult(End);
            }
            else if (User.IsInRole("GestorComercial"))
            {
                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {
                    var endereco = ProtheusInter.Sa1010s.FirstOrDefault(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico);

                    var Quants = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Bairro == endereco.A1Bairro).Count();

                    var Data = (SGID.Visitas.OrderByDescending(x => x.DataHora).FirstOrDefault(x => x.Medico == Medico && x.Empresa=="INTERMEDIC")?.DataHora) ?? DateTime.Now;
                    var End = new
                    {
                        Endereco = endereco.A1End,
                        Bairro = endereco.A1Bairro,
                        Quant = Quants,
                        Ultima = Data.ToString("yyyy-MM-dd")
                    };

                    return new JsonResult(End);
                }
                else
                {

                    var endereco = ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico);

                    var Quants = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Bairro == endereco.A1Bairro).Count();

                    var Data = (SGID.Visitas.OrderByDescending(x => x.DataHora).FirstOrDefault(x => x.Medico == Medico && x.Empresa == "DENUO")?.DataHora) ?? DateTime.Now;
                    var End = new
                    {
                        Endereco = endereco.A1End,
                        Bairro = endereco.A1Bairro,
                        Quant = Quants,
                        Ultima = Data.ToString("yyyy-MM-dd")
                    };

                    return new JsonResult(End);
                }
            }
            else
            {
                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {
                    var endereco = ProtheusInter.Sa1010s.FirstOrDefault(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico);

                    var Quants = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Bairro == endereco.A1Bairro).Count();

                    var Data = (SGID.Visitas.OrderByDescending(x => x.DataHora).FirstOrDefault(x => x.Medico == Medico && x.Empresa == "INTERMEDIC")?.DataHora) ?? DateTime.Now;
                    var End = new
                    {
                        Endereco = endereco.A1End,
                        Bairro = endereco.A1Bairro,
                        Quant = Quants,
                        Ultima = Data.ToString("yyyy-MM-dd")
                    };

                    return new JsonResult(End);
                }
                else
                {
                    var endereco = ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico);

                    var Quants = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Bairro == endereco.A1Bairro).Count();

                    var Data = (SGID.Visitas.OrderByDescending(x => x.DataHora).FirstOrDefault(x => x.Medico == Medico && x.Empresa == "DENUO")?.DataHora) ?? DateTime.Now;
                    var End = new
                    {
                        Endereco = endereco.A1End,
                        Bairro = endereco.A1Bairro,
                        Quant = Quants,
                        Ultima = Data.ToString("yyyy-MM-dd")
                    };

                    return new JsonResult(End);
                }
            }

        }

        public JsonResult OnGetMedico(string Bairro)
        {
            if (User.IsInRole("Admin"))
            {
                var medicos = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Bairro == Bairro).ToList();


                return new JsonResult(medicos);
            }
            else if (User.IsInRole("GestorComercial"))
            {
                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {
                    var medicos = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Bairro == Bairro).ToList();


                    return new JsonResult(medicos);
                }
                else
                {

                    var medicos = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Bairro == Bairro).ToList();


                    return new JsonResult(medicos);
                }
            }
            else
            {
                if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                {
                    var medicos = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Bairro == Bairro).ToList();


                    return new JsonResult(medicos);
                }
                else
                {
                    var medicos = ProtheusDenuo.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Bairro == Bairro).ToList();


                    return new JsonResult(medicos);
                }
            }
        }

        public class EventViewModel
        {
            public int Id { get; set; }
            public string Title { get; set; }
            public string Start { get; set; }
            public string End { get; set; }
            public bool AllDay { get; set; }
            //public string color { get; set; }
        }
    }
}
