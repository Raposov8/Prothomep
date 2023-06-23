using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;
using SGID.Models.Inter;
using SGID.Models.Patrimonio;
using System.Drawing;

namespace SGID.Pages.Formularios
{
    [Authorize]
    public class RelatorioPatrimonioTodosModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public List<Patrimonios> Relatorio { get; set; } = new List<Patrimonios>();
        public string Patrimonio { get; set; }
        public int? Total { get; set; }
        public int? Estoque { get; set; }
        public string Empresa { get; set; }
        public RelatorioPatrimonioTodosModel(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter,ApplicationDbContext sgid)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
            SGID = sgid;
        }

        public void OnGet(string id)
        {
            this.Empresa = id;
            Estoque = 0;
            Total = 0;
            this.Empresa = Empresa;
            if (Empresa == "01")
            {
                this.Patrimonio = Patrimonio;
                Relatorio = (from PA10 in ProtheusInter.Pa1010s
                             join PAC in ProtheusInter.Pac010s on PA10.Pa1Numage equals PAC.PacNumage
                             join SA10 in ProtheusInter.Sa1010s on new { Codigo = PAC.PacClient, Loja = PAC.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                             where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                             && PAC.DELET != "*" && SA10.DELET != "*"
                             select new Patrimonios
                             {
                                 Codigo = PA10.Pa1Codigo,
                                 DesPat = PA10.Pa1Despat,
                                 KitBas = PA10.Pa1Kitbas,
                                 DTCirurgia = $"{PAC.PacDtcir.Substring(6, 2)}/{PAC.PacDtcir.Substring(4, 2)}/{PAC.PacDtcir.Substring(0, 4)}",
                                 NMPac = PAC.PacNmpac,
                                 Status = PA10.Pa1Status == "N" ? "EM ESTOQUE" : PA10.Pa1Status == "S" ? "EM USO" : PA10.Pa1Status == "R" ? "AGUARDANDO REPOSIÇÃO" : PA10.Pa1Status == "Q" ? "INSPEÇÃO DE QUALIDADE" : PA10.Pa1Status == "E" ? "EMPENHO" : "",
                                 Cliente = PA10.Pa1Status == "S" ? SA10.A1Nreduz : "",
                                 Agend = PA10.Pa1Status == "S" ? PAC.PacNumage : "",
                                 QTDFALT = (from PA20 in ProtheusInter.Pa2010s
                                            join SB10 in ProtheusInter.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                                            where PA20.DELET != "*" && SB10.DELET != "*" &&
                                            PA20.Pa2Codigo == PA10.Pa1Codigo && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                                            select new
                                            {
                                                SB10.B1Cod
                                            }).Count(),
                                 Empresa = Empresa,
                                 Oper = PAC.PacTpoper,
                                 Valid = (from SB80 in ProtheusInter.Sb8010s
                                          join PA30 in ProtheusInter.Pa3010s on new { Filial = SB80.B8Filial, Local = SB80.B8Local, Produto = SB80.B8Produto, Lote = SB80.B8Lotectl, Codigo = PA10.Pa1Codigo } equals new { Filial = PA30.Pa3Filial, Local = PA30.Pa3Local, Produto = PA30.Pa3Produt, Lote = PA30.Pa3Lote, Codigo = PA30.Pa3Codigo }
                                          join SB10 in ProtheusInter.Sb1010s on PA30.Pa3Produt equals SB10.B1Cod
                                          where SB80.DELET != "*" && PA30.DELET != "*" && SB10.DELET != "*" && SB10.B1Tipo == "PA"
                                          select SB80.B8Dtvalid).FirstOrDefault(),
                                Filial = PA10.Pa1Filial
                             }
                             ).OrderBy(x => x.DesPat).ThenBy(x => x.Status).ToList();

                var query = ProtheusInter.Pa1010s.Where(x => x.DELET != "*" && x.Pa1Msblql != "1" && x.Pa1Status != "B")
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
                             && c.DELET != "*" && a.DELET != "*"
                             && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                             select new Patrimonios
                             {
                                 Codigo = PA10.Pa1Codigo,
                                 DesPat = PA10.Pa1Despat,
                                 KitBas = PA10.Pa1Kitbas,
                                 DTCirurgia = $"{c.PacDtcir.Substring(6, 2)}/{c.PacDtcir.Substring(4, 2)}/{c.PacDtcir.Substring(0, 4)}",
                                 NMPac = c.PacNmpac,
                                 Status = PA10.Pa1Status == "N" ? "EM ESTOQUE" : PA10.Pa1Status == "S" ? "EM USO" : PA10.Pa1Status == "R" ? "AGUARDANDO REPOSIÇÃO" : PA10.Pa1Status == "Q" ? "INSPEÇÃO DE QUALIDADE" : PA10.Pa1Status == "E" ? "EMPENHO" : "",
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
                                 Oper = c.PacTpoper,
                                 Valid = (from SB80 in ProtheusDenuo.Sb8010s
                                          join PA30 in ProtheusDenuo.Pa3010s on new { Filial = SB80.B8Filial, Local = SB80.B8Local, Produto = SB80.B8Produto, Lote = SB80.B8Lotectl, Codigo = PA10.Pa1Codigo } equals new { Filial = PA30.Pa3Filial, Local = PA30.Pa3Local, Produto = PA30.Pa3Produt, Lote = PA30.Pa3Lote, Codigo = PA30.Pa3Codigo }
                                          join SB10 in ProtheusDenuo.Sb1010s on PA30.Pa3Produt equals SB10.B1Cod
                                          where SB80.DELET != "*" && PA30.DELET != "*" && SB10.DELET != "*" && SB10.B1Tipo == "PA"
                                          select SB80.B8Dtvalid).FirstOrDefault(),
                                 Filial = PA10.Pa1Filial
                             }
                             ).OrderBy(x => x.DesPat).ThenBy(x=>x.Status).ToList();

                var query = ProtheusDenuo.Pa1010s.Where(x => x.DELET != "*" && x.Pa1Msblql != "1" && x.Pa1Status != "B")
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

            }
        }


        public IActionResult OnPostExport(string id)
        {
            try
            {
                this.Empresa = id;
                Estoque = 0;
                Total = 0;
                this.Empresa = Empresa;
                if (Empresa == "01")
                {
                    this.Patrimonio = Patrimonio;
                    Relatorio = (from PA10 in ProtheusInter.Pa1010s
                                 join PAC in ProtheusInter.Pac010s on PA10.Pa1Numage equals PAC.PacNumage
                                 join SA10 in ProtheusInter.Sa1010s on new { Codigo = PAC.PacClient, Loja = PAC.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                                 where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                                 && PAC.DELET != "*" && SA10.DELET != "*"
                                 select new Patrimonios
                                 {
                                     Codigo = PA10.Pa1Codigo,
                                     DesPat = PA10.Pa1Despat,
                                     KitBas = PA10.Pa1Kitbas,
                                     DTCirurgia = $"{PAC.PacDtcir.Substring(6, 2)}/{PAC.PacDtcir.Substring(4, 2)}/{PAC.PacDtcir.Substring(0, 4)}",
                                     NMPac = PAC.PacNmpac,
                                     Status = PA10.Pa1Status == "N" ? "EM ESTOQUE" : PA10.Pa1Status == "S" ? "EM USO" : PA10.Pa1Status == "R" ? "AGUARDANDO REPOSIÇÃO" : PA10.Pa1Status == "Q" ? "INSPEÇÃO DE QUALIDADE" : PA10.Pa1Status == "E" ? "EMPENHO" : "",
                                     Cliente = PA10.Pa1Status == "S" ? SA10.A1Nreduz : "",
                                     Agend = PA10.Pa1Status == "S" ? PAC.PacNumage : "",
                                     QTDFALT = (from PA20 in ProtheusInter.Pa2010s
                                                join SB10 in ProtheusInter.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                                                where PA20.DELET != "*" && SB10.DELET != "*" &&
                                                PA20.Pa2Codigo == PA10.Pa1Codigo && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                                                select new
                                                {
                                                    SB10.B1Cod
                                                }).Count(),
                                     Empresa = Empresa,
                                     Oper = PAC.PacTpoper,
                                     Valid = (from SB80 in ProtheusInter.Sb8010s
                                              join PA30 in ProtheusInter.Pa3010s on new { Filial = SB80.B8Filial, Local = SB80.B8Local, Produto = SB80.B8Produto, Lote = SB80.B8Lotectl, Codigo = PA10.Pa1Codigo } equals new { Filial = PA30.Pa3Filial, Local = PA30.Pa3Local, Produto = PA30.Pa3Produt, Lote = PA30.Pa3Lote, Codigo = PA30.Pa3Codigo }
                                              join SB10 in ProtheusInter.Sb1010s on PA30.Pa3Produt equals SB10.B1Cod
                                              where SB80.DELET != "*" && PA30.DELET != "*" && SB10.DELET != "*" && SB10.B1Tipo == "PA"
                                              select SB80.B8Dtvalid).FirstOrDefault()
                                 }
                                 ).ToList();

                    var query = ProtheusInter.Pa1010s.Where(x => x.DELET != "*" && x.Pa1Msblql != "1" && x.Pa1Status != "B")
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
                                 && c.DELET != "*" && a.DELET != "*"
                                 && ((int)(object)c.PacDtcir >= 20200701 || c.PacDtcir == null)
                                 select new Patrimonios
                                 {
                                     Codigo = PA10.Pa1Codigo,
                                     DesPat = PA10.Pa1Despat,
                                     KitBas = PA10.Pa1Kitbas,
                                     DTCirurgia = $"{c.PacDtcir.Substring(6, 2)}/{c.PacDtcir.Substring(4, 2)}/{c.PacDtcir.Substring(0, 4)}",
                                     NMPac = c.PacNmpac,
                                     Status = PA10.Pa1Status == "N" ? "EM ESTOQUE" : PA10.Pa1Status == "S" ? "EM USO" : PA10.Pa1Status == "R" ? "AGUARDANDO REPOSIÇÃO" : PA10.Pa1Status == "Q" ? "INSPEÇÃO DE QUALIDADE" : PA10.Pa1Status == "E" ? "EMPENHO" : "",
                                     Cliente = PA10.Pa1Status == "S" ? a.A1Nreduz : "",
                                     Agend = PA10.Pa1Status == "S" ? c.PacNumage : "",
                                     QTDFALT = (from PA20 in ProtheusDenuo.Pa2010s
                                                join SB10 in ProtheusDenuo.Sb1010s on PA20.Pa2Comp equals SB10.B1Cod
                                                where PA20.DELET != "*" && SB10.DELET != "*" &&
                                                PA20.Pa2Codigo == PA10.Pa1Codigo && PA20.Pa2Qtdkit > PA20.Pa2Qtdpat
                                                select new
                                                {
                                                    SB10.B1Cod
                                                }).Count(),
                                     Empresa = Empresa,
                                     Oper = c.PacTpoper,
                                     Valid = (from SB80 in ProtheusDenuo.Sb8010s
                                              join PA30 in ProtheusDenuo.Pa3010s on new { Filial = SB80.B8Filial, Local = SB80.B8Local, Produto = SB80.B8Produto, Lote = SB80.B8Lotectl, Codigo = PA10.Pa1Codigo } equals new { Filial = PA30.Pa3Filial, Local = PA30.Pa3Local, Produto = PA30.Pa3Produt, Lote = PA30.Pa3Lote, Codigo = PA30.Pa3Codigo }
                                              join SB10 in ProtheusDenuo.Sb1010s on PA30.Pa3Produt equals SB10.B1Cod
                                              where SB80.DELET != "*" && PA30.DELET != "*" && SB10.DELET != "*" && SB10.B1Tipo == "PA"
                                              select SB80.B8Dtvalid).FirstOrDefault()
                                 }
                                 ).OrderBy(x => x.DesPat).ToList();

                    var query = ProtheusDenuo.Pa1010s.Where(x => x.DELET != "*" && x.Pa1Msblql != "1" && x.Pa1Status != "B")
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

                }

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("DISPPATRT");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "DISPPATRT");

                sheet.Cells[1, 1].Value = $"Total de Patrimônios: {Total}"; 
                sheet.Cells[1, 6].Value = $"Em Estoque: {Estoque}";

                var format = sheet.Cells[1, 1, 1, 5];
                format.Merge = true;

                format = sheet.Cells[1, 6, 1, 10];
                format.Merge = true;

                sheet.Cells[2, 1].Value = "COMPLETO";
                sheet.Cells[2, 2].Value = "COD.";
                sheet.Cells[2, 3].Value = "PATRIMONIO";
                sheet.Cells[2, 4].Value = "STATUS";
                sheet.Cells[2, 5].Value = "CLIENTE";
                sheet.Cells[2, 6].Value = "PACIENTE";
                sheet.Cells[2, 7].Value = "DT. CIR.";
                sheet.Cells[2, 8].Value = "AGEND.";
                sheet.Cells[2, 9].Value = "OPER.";
                sheet.Cells[2, 9].Value = "VALID PA";

                int i = 3;

                Relatorio.ForEach(Pedido =>
                {
                    if (Pedido.QTDFALT > 0)
                    {
                        sheet.Cells[i, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[i, 1].Style.Fill.BackgroundColor.SetColor(Color.Red);
                    }
                    else
                    {
                        sheet.Cells[i, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[i, 1].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    }
                    sheet.Cells[i, 2].Value = Pedido.Codigo;
                    sheet.Cells[i, 3].Value = Pedido.DesPat;
                    sheet.Cells[i, 4].Value = Pedido.Status;
                    sheet.Cells[i, 5].Value = Pedido.Cliente;
                    sheet.Cells[i, 6].Value = Pedido.NMPac;
                    sheet.Cells[i, 7].Value = Pedido.DTCirurgia;
                    sheet.Cells[i, 8].Value = Pedido.Agend;
                    sheet.Cells[i, 9].Value = Pedido.Oper;
                    sheet.Cells[i, 10].Value = Pedido.Valid;

                    i++;
                });


                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DISPPATRT.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "DISPPATRT Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
