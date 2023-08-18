using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Models.Denuo;
using SGID.Models.Patrimonio;

namespace SGID.Pages.Formularios
{
    [Authorize]
    public class REQ09Model : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        public REQ09 Relatorio { get; set; }
        public string Empresa { get; set; }
        public DateTime Data { get; set; }
        public REQ09Model(TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter)
        {
            ProtheusDenuo = protheusDenuo;
            ProtheusInter = protheusInter;
        }

        public void OnGet(string id)
        {
            this.Empresa = id;
        }

        public IActionResult OnPost(string CirPrin, string CirCompl1,string CirCompl2,string Empresa)
        {
            this.Empresa = Empresa;
            Data = DateTime.Now;
            List<REQ09> lista = null;
            Relatorio = new REQ09();
            if (Empresa == "01")
            {
                var query = (from SUA010 in ProtheusInter.Sua010s
                             join PAC010 in ProtheusInter.Pac010s on new { Filial = SUA010.UaFilial, NumAge = SUA010.UaUnumage } equals new { Filial = PAC010.PacFilial, NumAge = PAC010.PacNumage }
                             join PAD010 in ProtheusInter.Pad010s on SUA010.UaUnumage equals PAD010.PadNumage
                             join SA1010 in ProtheusInter.Sa1010s on new { Cliente = SUA010.UaCliente, Loja = SUA010.UaLoja } equals new { Cliente = SA1010.A1Cod, Loja = SA1010.A1Loja }
                             join SA10 in ProtheusInter.Sa1010s on new { Convenio = SUA010.UaXConve, LojaCon = SUA010.UaXLjcon } equals new { Convenio = SA10.A1Cod, LojaCon = SA10.A1Loja } into sr
                             from x in sr.DefaultIfEmpty()
                             join SA3010 in ProtheusInter.Sa3010s on SUA010.UaVend equals SA3010.A3Cod
                             join SB1010 in ProtheusInter.Sb1010s on PAD010.PadProdut equals SB1010.B1Cod
                             where SUA010.DELET != "*" && PAC010.DELET != "*" && PAD010.DELET != "*" && SA1010.DELET != "*" && x.DELET != "*"
                             && SA3010.DELET != "*" && SB1010.DELET != "*" && SUA010.UaUnumage == CirPrin
                             orderby SUA010.UaUnumage, PAD010.PadProdut
                             select new REQ09
                             {
                                 NumProc = SUA010.UaNum,
                                 TipoCir = SUA010.UaUemerg == "S" ? "EM" : SUA010.UaTpcirur == "1" ? "EL" : SUA010.UaTpcirur == "2" ? "EM" : "N/D",
                                 Codcli = SUA010.UaCliente,
                                 LojaCli = SUA010.UaLoja,
                                 Cliente = SA1010.A1Nreduz,
                                 DtCir = $"{SUA010.UaXDtcir.Substring(6, 2)}/{SUA010.UaXDtcir.Substring(4, 2)}/{SUA010.UaXDtcir.Substring(0, 4)}",
                                 HrCir = SUA010.UaXHrcir,
                                 Paciente = SUA010.UaXNmpac,
                                 Convenio = $"{SUA010.UaXConve}-{SUA010.UaXLjcon} {x.A1Nome}",
                                 Medico = $"{SUA010.UaXMedic}-{SUA010.UaXNmmed}",
                                 Vendedor = $"{SUA010.UaVend}-{SA3010.A3Nome}",
                                 Material = $"{SB1010.B1Desc.Trim()} ({PAD010.PadPatrim})",
                                 Obs = System.Text.Encoding.UTF8.GetString(PAC010.PacObs),
                                 Cirurgia = SUA010.UaUnumage,

                                 Produto = PAD010.PadProdut,
                                 Patrimonio = PAD010.PadPatrim
                             });

                if(!string.IsNullOrEmpty(CirPrin) && !string.IsNullOrWhiteSpace(CirPrin))
                {

                    if(!string.IsNullOrEmpty(CirCompl1)&& !string.IsNullOrWhiteSpace(CirCompl1))
                    {
                        if (!string.IsNullOrEmpty(CirCompl2) && !string.IsNullOrWhiteSpace(CirCompl2))
                        {
                            lista = query.Where(x => x.Cirurgia == CirPrin || x.Cirurgia == CirCompl1 || x.Cirurgia == CirCompl2).OrderBy(x => x.Cirurgia).ThenBy(x => x.Produto).ToList();
                        }
                        else
                        {
                            lista = query.Where(x => x.Cirurgia == CirPrin || x.Cirurgia == CirCompl1).OrderBy(x => x.Cirurgia).ThenBy(x => x.Produto).ToList();
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(CirCompl2) && !string.IsNullOrWhiteSpace(CirCompl2))
                        {
                            lista = query.Where(x => x.Cirurgia == CirPrin || x.Cirurgia == CirCompl2).OrderBy(x => x.Cirurgia).ThenBy(x => x.Produto).ToList();
                        }
                        else
                        {
                            lista = query.Where(x => x.Cirurgia == CirPrin).OrderBy(x=>x.Cirurgia).ThenBy(x=>x.Produto).ToList();
                        }
                    }
                }
            }
            else
            {

                var query = (from SUA010 in ProtheusDenuo.Sua010s
                             join PAC010 in ProtheusDenuo.Pac010s on new { Filial = SUA010.UaFilial, NumAge = SUA010.UaUnumage } equals new { Filial = PAC010.PacFilial, NumAge = PAC010.PacNumage }
                             join PAD010 in ProtheusDenuo.Pad010s on SUA010.UaUnumage equals PAD010.PadNumage
                             join SA1010 in ProtheusDenuo.Sa1010s on new { Cliente = SUA010.UaCliente, Loja = SUA010.UaLoja } equals new { Cliente = SA1010.A1Cod, Loja = SA1010.A1Loja }
                             join SA10 in ProtheusDenuo.Sa1010s on new { Convenio = SUA010.UaXConve, LojaCon = SUA010.UaXLjcon } equals new { Convenio = SA10.A1Cod, LojaCon = SA10.A1Loja } into sr

                             from x in sr.DefaultIfEmpty()
                             join SA3010 in ProtheusDenuo.Sa3010s on SUA010.UaVend equals SA3010.A3Cod
                             join SB1010 in ProtheusDenuo.Sb1010s on PAD010.PadProdut equals SB1010.B1Cod
                             where SUA010.DELET != "*" && PAC010.DELET != "*" && PAD010.DELET != "*" && SA1010.DELET != "*" && x.DELET != "*"
                             && SA3010.DELET != "*" && SB1010.DELET != "*" && SUA010.UaUnumage == CirPrin
                             orderby SUA010.UaUnumage, PAD010.PadProdut
                             select new REQ09
                             {
                                 NumProc = SUA010.UaNum,
                                 TipoCir = SUA010.UaUemerg == "S" ? "EM" : SUA010.UaTpcirur == "1" ? "EL" : SUA010.UaTpcirur == "2" ? "EM" : "N/D",
                                 Codcli = SUA010.UaCliente,
                                 LojaCli = SUA010.UaLoja,
                                 Cliente = SA1010.A1Nreduz,
                                 DtCir = $"{SUA010.UaXDtcir.Substring(6, 2)}/{SUA010.UaXDtcir.Substring(4, 2)}/{SUA010.UaXDtcir.Substring(0, 4)}",
                                 HrCir = SUA010.UaXHrcir,
                                 Paciente = SUA010.UaXNmpac,
                                 Convenio = $"{SUA010.UaXConve}-{SUA010.UaXLjcon} {x.A1Nome}",
                                 Medico = $"{SUA010.UaXMedic}-{SUA010.UaXNmmed}",
                                 Vendedor = $"{SUA010.UaVend}-{SA3010.A3Nome}",
                                 Material = $"{SB1010.B1Desc.Trim()} ({PAD010.PadPatrim})",
                                 Obs = System.Text.Encoding.UTF8.GetString(PAC010.PacObs),
                                 Cirurgia = SUA010.UaUnumage,

                                 Produto = PAD010.PadProdut,
                                 Patrimonio = PAD010.PadPatrim
                             });

                if (!string.IsNullOrEmpty(CirPrin) && !string.IsNullOrWhiteSpace(CirPrin))
                {

                    if (!string.IsNullOrEmpty(CirCompl1) && !string.IsNullOrWhiteSpace(CirCompl1))
                    {
                        if (!string.IsNullOrEmpty(CirCompl2) && !string.IsNullOrWhiteSpace(CirCompl2))
                        {
                            lista = query.Where(x => x.Cirurgia == CirPrin || x.Cirurgia == CirCompl1 || x.Cirurgia == CirCompl2).OrderBy(x => x.Cirurgia).ThenBy(x => x.Produto).ToList();
                        }
                        else
                        {
                            lista = query.Where(x => x.Cirurgia == CirPrin || x.Cirurgia == CirCompl1).OrderBy(x => x.Cirurgia).ThenBy(x => x.Produto).ToList();
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(CirCompl2) && !string.IsNullOrWhiteSpace(CirCompl2))
                        {
                            lista = query.Where(x => x.Cirurgia == CirPrin || x.Cirurgia == CirCompl2).OrderBy(x => x.Cirurgia).ThenBy(x => x.Produto).ToList();
                        }
                        else
                        {
                            lista = query.Where(x => x.Cirurgia == CirPrin).OrderBy(x => x.Cirurgia).ThenBy(x => x.Produto).ToList();
                        }
                    }
                }
            }

            if (lista != null)
            {
                Relatorio = lista.First();

                string CirurgiaAnt = Relatorio.Cirurgia;
                string ProdutoAnt = Relatorio.Produto;
                string PatrimonioAnt = Relatorio.Patrimonio;

                if (Empresa == "01")
                {
                    var query = (from PAJ010 in ProtheusInter.Paj010s
                                 join PAI010 in ProtheusInter.Pai010s on PAJ010.PajPrccir equals PAI010.PaiGrpprc
                                 join PAH010 in ProtheusInter.Pah010s on PAJ010.PajCodins equals PAH010.PahCodins into sr
                                 from c in sr.DefaultIfEmpty()
                                 where PAJ010.DELET != "*" && PAI010.DELET != "*" && c.DELET != "*"
                                 && PAJ010.PajProces == Relatorio.NumProc && c.PahNome != "XXXXX"
                                 select new
                                 {
                                     c.PahNome,
                                     PAI010.PaiDescr
                                 }).FirstOrDefault();

                    Relatorio.NomInst = query?.PahNome;
                    Relatorio.Procedimento = query?.PaiDescr;
                }
                else
                {
                   var query = (from PAJ010 in ProtheusDenuo.Paj010s
                                 join PAI010 in ProtheusDenuo.Pai010s on PAJ010.PajPrccir equals PAI010.PaiGrpprc
                                 join PAH010 in ProtheusDenuo.Pah010s on PAJ010.PajCodins equals PAH010.PahCodins into sr
                                 from c in sr.DefaultIfEmpty()
                                 where PAJ010.DELET != "*" && PAI010.DELET != "*" && c.DELET != "*"
                                 && PAJ010.PajProces == Relatorio.NumProc && c.PahNome != "XXXXX"
                                 select new
                                 {
                                     c.PahNome,
                                     PAI010.PaiDescr
                                 }).FirstOrDefault();

                    Relatorio.NomInst = query?.PahNome;
                    Relatorio.Procedimento = query?.PaiDescr;
                }

                

                lista.ForEach(x =>
                {

                    if (CirurgiaAnt != x.Cirurgia)
                    {
                        Relatorio.Obs += $" | {x.Obs}";
                        Relatorio.Cirurgia += $" | {x.Cirurgia}";
                        Relatorio.NumProc += $" | {x.NumProc}";
                        CirurgiaAnt = x.Cirurgia;

                        if (Empresa == "01")
                        {
                            var query = (from PAJ010 in ProtheusInter.Paj010s
                                         join PAI010 in ProtheusInter.Pai010s on PAJ010.PajPrccir equals PAI010.PaiGrpprc
                                         join PAH010 in ProtheusInter.Pah010s on PAJ010.PajCodins equals PAH010.PahCodins into sr
                                         from c in sr.DefaultIfEmpty()
                                         where PAJ010.DELET != "*" && PAI010.DELET != "*" && c.DELET != "*"
                                         && PAJ010.PajProces == x.NumProc && c.PahNome != "XXXXX"
                                         select new
                                         {
                                             c.PahNome,
                                             PAI010.PaiDescr
                                         }).FirstOrDefault();

                            if (query != null)
                            {
                                Relatorio.NomInst += $" | {query?.PahNome}";
                                Relatorio.Procedimento += $" | {query?.PaiDescr}";
                            }
                        }
                        else
                        {
                            var query = (from PAJ010 in ProtheusDenuo.Paj010s
                                         join PAI010 in ProtheusDenuo.Pai010s on PAJ010.PajPrccir equals PAI010.PaiGrpprc
                                         join PAH010 in ProtheusDenuo.Pah010s on PAJ010.PajCodins equals PAH010.PahCodins into sr
                                         from c in sr.DefaultIfEmpty()
                                         where PAJ010.DELET != "*" && PAI010.DELET != "*" && c.DELET != "*"
                                         && PAJ010.PajProces == x.NumProc && c.PahNome != "XXXXX"
                                         select new
                                         {
                                             c.PahNome,
                                             PAI010.PaiDescr
                                         }).FirstOrDefault();

                            if (query != null)
                            {
                                Relatorio.NomInst += $" | {query?.PahNome}";
                                Relatorio.Procedimento += $" | {query?.PaiDescr}";
                            }
                        }
                    }
                    else
                    {
                        if (ProdutoAnt != x.Produto || PatrimonioAnt != x.Patrimonio)
                        {
                            Relatorio.Material += $" | {x.Material}";
                            ProdutoAnt = x.Produto;
                            PatrimonioAnt = x.Patrimonio;
                        }
                    }
                });
            }

            return Page();
        }
    }
}
