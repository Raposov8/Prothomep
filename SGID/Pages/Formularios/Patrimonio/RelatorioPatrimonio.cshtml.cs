using DocumentFormat.OpenXml.Office2010.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Models.Controladoria;
using SGID.Models.Denuo;
using SGID.Models.Patrimonio;

namespace SGID.Pages.Formularios.Patrimonio
{
    [Authorize]
    public class RelatorioPatrimonioModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        public List<Patrimonios> Relatorio { get; set; } = new List<Patrimonios>();
        public List<string> Patrimonios { get; set; } = new List<string>();
        public string Patrimonio { get; set; }
        public int? Total { get; set; }
        public int? Estoque { get; set; } 
        public string Empresa { get; set; }
        public RelatorioPatrimonioModel(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
        }

        public void OnGet(string id)
        {
            this.Empresa = id;
            if (id == "01") Patrimonios = ProtheusInter.Pa1010s.Where(x => x.DELET != "*" && x.Pa1Msblql != "1" && x.Pa1Status != "B").Select(x => x.Pa1Despat).Distinct().ToList();
            else Patrimonios = (from PA10 in ProtheusDenuo.Pa1010s
                                join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                from c in sr.DefaultIfEmpty()
                                join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                                from a in st.DefaultIfEmpty()
                                where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                                && c.DELET != "*" && a.DELET != "*" && PA10.DELET != "*"
                                && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                                select PA10.Pa1Despat ).Distinct().ToList();
        }

        public IActionResult OnGetConfirmar(string Patrimonio, string Empresa)
        {
            Estoque = 0;
            Total = 0;
            this.Empresa = Empresa;
            if (Empresa == "01")
            {
                this.Patrimonio = Patrimonio.Trim();
                Relatorio = (from PA10 in ProtheusInter.Pa1010s
                             join PAC in ProtheusInter.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                             from c in sr.DefaultIfEmpty()
                             join SA10 in ProtheusInter.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                             from a in st.DefaultIfEmpty()
                             where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                             && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == Patrimonio
                             select new Patrimonios
                             {
                                 Codigo = PA10.Pa1Codigo,
                                 DesPat = PA10.Pa1Despat,
                                 KitBas = PA10.Pa1Kitbas,
                                 DTCirurgia = PA10.Pa1Status == "S" ? $"{c.PacDtcir.Substring(6, 2)}/{c.PacDtcir.Substring(4, 2)}/{c.PacDtcir.Substring(0, 4)}" : "",
                                 NMPac = PA10.Pa1Status == "S" ? c.PacNmpac : "",
                                 Status = PA10.Pa1Status == "N" ? "EM ESTOQUE" : PA10.Pa1Status == "S" ? "EM USO" : PA10.Pa1Status == "R" ? "AGUARDANDO REPOSI플O" : PA10.Pa1Status == "Q" ? "INSPE플O DE QUALIDADE" : PA10.Pa1Status == "E" ? "EMPENHO" : "",
                                 Cliente = PA10.Pa1Status == "S" ? a.A1Nreduz : "",
                                 Agend = PA10.Pa1Status == "S" ? c.PacNumage : "",
                                 QTDFALT = (from PA20 in ProtheusInter.Pa2010s
                                            join SB10 in ProtheusInter.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                                            where PA20.DELET != "*" && SB10.DELET != "*" &&
                                            PA20.Pa2Codigo == PA10.Pa1Codigo && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                                            select new
                                            {
                                                SB10.B1Cod
                                            }).Count(),
                                 Empresa = Empresa,
                                 Filial = PA10.Pa1Filial
                             }
                             ).OrderBy(x => x.Status).ToList();

                var query = ProtheusInter.Pa1010s.Where(x => x.DELET != "*" && x.Pa1Msblql != "1" && x.Pa1Status != "B" && x.Pa1Despat == Patrimonio)
                    .Select(x => new
                    {
                        Status = x.Pa1Status == "N" ? "E" : "O",
                        Codigo = x.Pa1Codigo
                    })
                    .GroupBy(x => x.Status).Select(x => new
                    {
                        Status = x.Key,
                        QTDCOD = x.Count()
                    }).ToList();

                query.ForEach(x => {

                    if (x.Status == "E")
                    {
                        Estoque += x.QTDCOD;
                        Total += x.QTDCOD;
                    }
                    else
                    {
                        Total += x.QTDCOD;
                    }
                });

                Patrimonios = ProtheusInter.Pa1010s.Where(x => x.DELET != "*" && x.Pa1Msblql != "1" && x.Pa1Status != "B").Select(x => x.Pa1Despat).Distinct().ToList();

            }
            else
            {
                this.Patrimonio = Patrimonio.Trim();
                Relatorio = (from PA10 in ProtheusDenuo.Pa1010s
                             join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                             from c in sr.DefaultIfEmpty()
                             join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                             from a in st.DefaultIfEmpty()
                             where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                             && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == Patrimonio
                             && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                             select new Patrimonios
                             {
                                 Codigo = PA10.Pa1Codigo,
                                 DesPat = PA10.Pa1Despat,
                                 KitBas = PA10.Pa1Kitbas,
                                 DTCirurgia = PA10.Pa1Status == "S" ? $"{c.PacDtcir.Substring(6, 2)}/{c.PacDtcir.Substring(4, 2)}/{c.PacDtcir.Substring(0, 4)}" : "",
                                 NMPac = PA10.Pa1Status == "S" ? c.PacNmpac : "",
                                 Status = PA10.Pa1Status == "N" ? "EM ESTOQUE" : PA10.Pa1Status == "S" ? "EM USO" : PA10.Pa1Status == "R" ? "AGUARDANDO REPOSI플O" : PA10.Pa1Status == "Q" ? "INSPE플O DE QUALIDADE" : PA10.Pa1Status == "E" ? "EMPENHO" : "",
                                 Cliente = PA10.Pa1Status == "S" ? a.A1Nreduz : "",
                                 Agend = PA10.Pa1Status == "S" ? c.PacNumage : "",
                                 QTDFALT = (from PA20 in ProtheusDenuo.Pa2010s
                                            join SB10 in ProtheusDenuo.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                                            where PA20.DELET != "*" && SB10.DELET != "*" &&
                                            PA20.Pa2Codigo == PA10.Pa1Codigo && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                                            && PA20.Pa2Filial == PA10.Pa1Filial
                                            select new
                                            {
                                                SB10.B1Cod
                                            }).Count(),
                                 Empresa = Empresa,
                                 Filial = PA10.Pa1Filial
                             }
                             ).OrderBy(x => x.Status).ToList();

                var query = (from PA10 in ProtheusDenuo.Pa1010s
                             join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                             from c in sr.DefaultIfEmpty()
                             join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                             from a in st.DefaultIfEmpty()
                             where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                             && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == Patrimonio
                             && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                             select new
                             {
                                 Status = PA10.Pa1Status == "N" ? "E" : "O",
                                 Codigo = PA10.Pa1Codigo
                             }
                             )
                             .Select(x => new
                             {
                                 Status = x.Status,
                                 QTDCOD = 1
                             }).ToList();

                query.ForEach(x => {

                    if (x.Status == "E")
                    {
                        Estoque += x.QTDCOD;
                        Total += x.QTDCOD;
                    }
                    else
                    {
                        Total += x.QTDCOD;
                    }
                });

                Patrimonios = (from PA10 in ProtheusDenuo.Pa1010s
                               join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                               from c in sr.DefaultIfEmpty()
                               join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                               from a in st.DefaultIfEmpty()
                               where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                               && c.DELET != "*" && a.DELET != "*" && PA10.DELET != "*"
                               && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                               select PA10.Pa1Despat
                                   ).Distinct().ToList();
            }

            return Page();
        }

        public IActionResult OnPost(string Patrimonio,string Empresa)
        {
            Estoque = 0;
            Total = 0;
            this.Empresa = Empresa;
            if(Empresa == "01")
            {
                this.Patrimonio = Patrimonio.Trim(); 
                Relatorio = (from PA10 in ProtheusInter.Pa1010s
                             join PAC in ProtheusInter.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                             from c in sr.DefaultIfEmpty()
                             join SA10 in ProtheusInter.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                             from a in st.DefaultIfEmpty()
                             where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                             && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == Patrimonio
                             select new Patrimonios
                             {
                                 Codigo = PA10.Pa1Codigo,
                                 DesPat = PA10.Pa1Despat,
                                 KitBas = PA10.Pa1Kitbas,
                                 DTCirurgia = PA10.Pa1Status == "S" ? $"{c.PacDtcir.Substring(6, 2)}/{c.PacDtcir.Substring(4, 2)}/{c.PacDtcir.Substring(0, 4)}" : "",
                                 NMPac = PA10.Pa1Status == "S" ? c.PacNmpac : "",
                                 Status = PA10.Pa1Status == "N" ? "EM ESTOQUE" : PA10.Pa1Status == "S" ? "EM USO" : PA10.Pa1Status == "R" ? "AGUARDANDO REPOSI플O" : PA10.Pa1Status == "Q" ? "INSPE플O DE QUALIDADE" : PA10.Pa1Status == "E" ? "EMPENHO" : "",
                                 Cliente = PA10.Pa1Status == "S" ? a.A1Nreduz : "",
                                 Agend = PA10.Pa1Status == "S" ? c.PacNumage : "",
                                 QTDFALT = (from PA20 in ProtheusInter.Pa2010s
                                            join SB10 in ProtheusInter.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                                            where PA20.DELET != "*" && SB10.DELET != "*" &&
                                            PA20.Pa2Codigo == PA10.Pa1Codigo && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                                            select new
                                            {
                                                SB10.B1Cod
                                            }).Count(),
                                 Empresa = Empresa,
                                 Filial = PA10.Pa1Filial
                             }
                             ).OrderBy(x=> x.Status).ToList();

                var query = ProtheusInter.Pa1010s.Where(x => x.DELET != "*" && x.Pa1Msblql != "1" && x.Pa1Status != "B" && x.Pa1Despat == Patrimonio)
                    .Select(x => new
                    {
                        Status = x.Pa1Status == "N" ? "E" : "O",
                        Codigo = x.Pa1Codigo
                    })
                    .GroupBy(x => x.Status).Select(x => new
                    {
                        Status = x.Key,
                        QTDCOD = x.Count()
                    }).ToList();

                query.ForEach(x => {

                    if (x.Status == "E")
                    {
                        Estoque += x.QTDCOD;
                        Total += x.QTDCOD;
                    }
                    else
                    {
                        Total += x.QTDCOD;
                    }
                });

                Patrimonios = ProtheusInter.Pa1010s.Where(x => x.DELET != "*" && x.Pa1Msblql != "1" && x.Pa1Status != "B").Select(x => x.Pa1Despat).Distinct().ToList();
                
            }
            else
            {
                this.Patrimonio = Patrimonio;
                Relatorio = (from PA10 in ProtheusDenuo.Pa1010s
                             join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                             from c in sr.DefaultIfEmpty()
                             join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                             from a in st.DefaultIfEmpty()
                             where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                             && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == Patrimonio
                             && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                             select new Patrimonios
                             {
                                 Codigo = PA10.Pa1Codigo,
                                 DesPat = PA10.Pa1Despat,
                                 KitBas = PA10.Pa1Kitbas,
                                 DTCirurgia = PA10.Pa1Status == "S" ? $"{c.PacDtcir.Substring(6, 2)}/{c.PacDtcir.Substring(4, 2)}/{c.PacDtcir.Substring(0, 4)}":"",
                                 NMPac = PA10.Pa1Status == "S" ? c.PacNmpac:"",
                                 Status = PA10.Pa1Status == "N" ? "EM ESTOQUE" : PA10.Pa1Status == "S" ? "EM USO" : PA10.Pa1Status == "R" ? "AGUARDANDO REPOSI플O" : PA10.Pa1Status == "Q" ? "INSPE플O DE QUALIDADE" : PA10.Pa1Status == "E" ? "EMPENHO" : "",
                                 Cliente = PA10.Pa1Status == "S" ? a.A1Nreduz : "",
                                 Agend = PA10.Pa1Status == "S" ? c.PacNumage : "",
                                 QTDFALT = (from PA20 in ProtheusDenuo.Pa2010s
                                            join SB10 in ProtheusDenuo.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                                            where PA20.DELET != "*" && SB10.DELET != "*" &&
                                            PA20.Pa2Codigo == PA10.Pa1Codigo && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                                            && PA20.Pa2Filial == PA10.Pa1Filial
                                            select new
                                            {
                                                SB10.B1Cod
                                            }).Count(),
                                 Empresa = Empresa,
                                 Filial = PA10.Pa1Filial
                             }
                             ).OrderBy(x => x.Status).ToList();

                var query = (from PA10 in ProtheusDenuo.Pa1010s
                             join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                             from c in sr.DefaultIfEmpty()
                             join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                             from a in st.DefaultIfEmpty()
                             where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                             && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == Patrimonio
                             && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                             select new
                             {
                                 Status = PA10.Pa1Status == "N" ? "E" : "O",
                                 Codigo = PA10.Pa1Codigo
                             }
                             )
                             .Select(x => new
                             {
                                 Status = x.Status,
                                 QTDCOD = 1
                             }).ToList();

                query.ForEach(x => {

                    if (x.Status == "E")
                    {
                        Estoque += x.QTDCOD;
                        Total += x.QTDCOD;
                    }
                    else
                    {
                        Total += x.QTDCOD;
                    }
                });

                Patrimonios = (from PA10 in ProtheusDenuo.Pa1010s
                               join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                               from c in sr.DefaultIfEmpty()
                               join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                               from a in st.DefaultIfEmpty()
                               where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                               && c.DELET != "*" && a.DELET != "*" && PA10.DELET != "*"
                               && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                               select PA10.Pa1Despat
                                   ).Distinct().ToList();
            }

            return Page();
        }
    }
}
