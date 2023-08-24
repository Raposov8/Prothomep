using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;
using SGID.Models.DTO;
using System.Globalization;

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
        public List<string> Locais { get; set; } = new List<string> { "HOSPITAL","CLINICA","ESCRITORIO","CONSULTORIO","CASA","HOTEL","LOCAL PUBLICO" };

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

            if (User.IsInRole("Admin") || User.IsInRole("Diretoria"))
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
                if (User.IsInRole("Admin") || User.IsInRole("Diretoria"))
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
                    var cor = x.Status == 0? "blue" : x.Status == 1 ? "green" : x.Status == 2 ? "yellow" : "red";

                    events.Add(new EventViewModel()
                    {
                        Id = x.Id,
                        Title = x.Medico,
                        Start = $"{x.DataHora.ToString("o", CultureInfo.InvariantCulture)}",
                        color = cor,
                        textColor = cor == "yellow" ? "#20232a" : "#ffffff",
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
                    .Select(x => new VisitaDTO
                    {
                        Id = x.Id,
                        DataCriacao = x.DataHora,
                        DataHora = x.DataHora.ToString("dd/MM/yyyy HH:mm"),
                        DataUltima = x.DataUltima != null ? x.DataUltima.Value.ToString("dd/MM/yyyy") :"Não há registro",
                        Local = x.Local,
                        Assunto = x.Assunto,
                        Observacao = x.Observacao,
                        Endereco = x.Endereco,
                        Medico = x.Medico,
                        Motvisita = x.Motvisita ?? "",
                        Bairro = x.Bairro,
                        ResuVi = x.ResumoVisita,
                        Vendedor = x.Vendedor.Replace("."," "),
                        Status = x.Status,
                       DataProxima = x.DataHora.ToString("yyyy-MM-dd"),
                       UltimaResp1 = "",
                       UltimaResp2 = "",
                       Email = x.Email ?? "",
                       Telefone = x.Telefone ?? ""
                    }).FirstOrDefault(x=> x.Id == id);


                var ultimas = SGID.Visitas.Where(x => x.Medico == visita.Medico && x.DataHora < visita.DataCriacao && x.Status == 1).OrderByDescending(x => x.DataHora).Take(2).ToList();

                foreach (var data in ultimas)
                {
                    if(visita.UltimaResp1 == "")
                    {
                        visita.UltimaResp1 = data.ResumoVisita ?? "Não Há Resumo";
                        visita.DataResp1 = data.DataHora.ToString("dd/MM/yyyy HH:mm");
                    }
                    else
                    {
                        visita.UltimaResp2 = data.ResumoVisita ?? "Não Há Resumo";
                        visita.DataResp2 = data.DataHora.ToString("dd/MM/yyyy HH:mm");
                    }
                }


                return new JsonResult(visita);
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Visitas Details", user);
            }

            return new JsonResult("");
        }

        public JsonResult OnPost(string Medico, string Local, DateTime Data, DateTime DataUltima, string Assunto,
            string Endereco,string Bairro,string Email,string Telefone, string MotVis, string Obs,string ResuVisi)
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
                    Telefone = Telefone,
                    Email = Email,
                    Assunto = Assunto,
                    Endereco = Endereco,
                    Bairro = Bairro,
                    Motvisita = MotVis,
                    Observacao = Obs,
                    ResumoVisita = ResuVisi,
                    Vendedor = User.Identity.Name.Split("@")[0].ToUpper(),
                    Empresa = User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR" ? "INTERMEDIC":"DENUO"
                };

                if(!SGID.Visitas.Any(x=> x.DataHora == Visita.DataHora && x.Medico == Visita.Medico))
                {
                    SGID.Visitas.Add(Visita);
                    SGID.SaveChanges();

                    if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                    {
                        if (ProtheusInter.Sa1010s.Any(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico))
                        {
                            if (SGID.VisitaClientes.Any(x => x.Medico == Medico && x.Empresa == "INTERMEDIC"))
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
                    }
                    else
                    {
                        if (ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico) == null)
                        {
                            if (SGID.VisitaClientes.Any(x => x.Medico == Medico && x.Empresa == "DENUO"))
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
                    }
                }

                return new JsonResult("Sucesso");

            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovaVisita", user);
                return new JsonResult("");
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
                visita.Status = 1;
                

                SGID.Visitas.Update(visita);
                SGID.SaveChanges();

                return new JsonResult("");

            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Post Resumo Visita", user);
                return new JsonResult("");
            }
        }

        public JsonResult OnGetEndereco(string Medico)
        {
            try
            {
                if (User.IsInRole("Admin") || User.IsInRole("Diretoria"))
                {
                    var endereco = ProtheusInter.Sa1010s.FirstOrDefault(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico);

                    var Quants = ProtheusInter.Sa1010s.Where(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Bairro == endereco.A1Bairro).Count();

                    var Data = (SGID.Visitas.OrderByDescending(x => x.DataHora).FirstOrDefault(x => x.Medico == Medico && x.Empresa == "INTERMEDIC")?.DataHora) ?? DateTime.Now;

                    var End = new
                    {
                        Endereco = endereco.A1End,
                        Bairro = endereco.A1Bairro,
                        Quant = Quants,
                        Ultima = Data.ToString("yyyy-MM-dd"),
                        Telefone = $"{endereco.A1Ddd}{endereco.A1Tel}",
                        Email = endereco.A1Email
                    };

                    return new JsonResult(End);
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
                            Ultima = Data.ToString("yyyy-MM-dd"),
                            Telefone = $"{endereco.A1Ddd}{endereco.A1Tel}",
                            Email = endereco.A1Email
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
                            Ultima = Data.ToString("yyyy-MM-dd"),
                            Telefone = $"{endereco.A1Ddd}{endereco.A1Tel}",
                            Email = endereco.A1Email
                        };

                        return new JsonResult(End);
                    }
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Visita Endereco", user);
                return new JsonResult("");
            }


        }

        public JsonResult OnGetMedico(string Bairro)
        {
            if (User.IsInRole("Admin") || User.IsInRole("Diretoria"))
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

        public JsonResult OnPostCancelar(int Id,string MotivoRej)
        {
            try
            {
                var visita = SGID.Visitas.FirstOrDefault(x => x.Id == Id);

                visita.Status = 3;
                visita.MotivoRej = MotivoRej;

                SGID.Visitas.Update(visita);
                SGID.SaveChanges();

                return new JsonResult("");

            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Cancelar Visita", user);
                return new JsonResult("");
            }
        }

        public JsonResult OnPostRemarcar(int id,string Medico, string Local, DateTime Data, DateTime DataUltima, string Assunto,
            string Endereco, string Bairro, string MotVis, string Obs, string ResuVisi)
        {
            try
            {
                var VisitaRemark = SGID.Visitas.FirstOrDefault(x => x.Id == id);

                VisitaRemark.Status = 2;

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
                    Vendedor = VisitaRemark.Vendedor,
                    Empresa = VisitaRemark.Empresa
                };

                if (!SGID.Visitas.Any(x => x.DataHora == Visita.DataHora && x.Medico == Visita.Medico))
                {
                    SGID.Visitas.Add(Visita);
                    SGID.SaveChanges();


                    VisitaRemark.IdRemarcar = Visita.Id;
                    SGID.Visitas.Update(VisitaRemark);
                    SGID.SaveChanges();


                    if (User.Identity.Name.Split("@")[1].ToUpper() == "INTERMEDIC.COM.BR")
                    {
                        if (ProtheusInter.Sa1010s.Any(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico))
                        {
                            if (SGID.VisitaClientes.Any(x => x.Medico == Medico && x.Empresa == "INTERMEDIC"))
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
                    }
                    else
                    {
                        if (ProtheusDenuo.Sa1010s.FirstOrDefault(x => x.DELET != "*" && x.A1Clinter == "M" && x.A1Msblql != "1" && !string.IsNullOrWhiteSpace(x.A1Vend) && x.A1Nome == Medico) == null)
                        {
                            if (SGID.VisitaClientes.Any(x => x.Medico == Medico && x.Empresa == "DENUO"))
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
                    }
                }

                return new JsonResult("Sucesso");

            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NovaVisita", user);
            }

            return new JsonResult("");
        }

        public class EventViewModel
        {
            public int Id { get; set; }
            public string Title { get; set; }
            public string Start { get; set; }
            public string End { get; set; }
            public bool AllDay { get; set; }
            public string color { get; set; }
            public string textColor { get; set; }
        }
    }
}
