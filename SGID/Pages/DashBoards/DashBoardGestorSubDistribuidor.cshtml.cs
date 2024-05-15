using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Account.RH;
using SGID.Models.Denuo;
using SGID.Models.Diretoria;
using SGID.Models.Financeiro;
using SGID.Models.Inter;

namespace SGID.Pages.DashBoards
{
    public class DashBoardGestorSubDistribuidorModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }

        public List<ValoresEmAberto> ValoresEmAberto { get; set; } = new List<ValoresEmAberto>();
        public List<ValoresEmAberto> Faturamento { get; set; } = new List<ValoresEmAberto>();
        public List<ValoresEmAberto> ClientesDenuo { get; set; } = new List<ValoresEmAberto>();
        public List<ValoresEmAberto> ClientesInter { get; set; } = new List<ValoresEmAberto>();
        public List<ValoresEmAberto> FabricanteInter { get; set; } = new List<ValoresEmAberto>();
        public List<ValoresEmAberto> FabricanteDenuo { get; set; } = new List<ValoresEmAberto>();
        public string Ano { get; set; }

        public DashBoardGestorSubDistribuidorModel(ApplicationDbContext sgid, TOTVSDENUOContext denuo, TOTVSINTERContext inter)
        {
            SGID = sgid;
            ProtheusDenuo = denuo;
            ProtheusInter = inter;

        }

        public void OnGet()
        {
            Ano = DateTime.Now.Year.ToString();
        }


        public JsonResult OnPostFaturados(string Mes, string Ano)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                string[] CF = new string[] { "5551", "6551", "6107", "6109" };
                string DataInicio = $"{Ano}0101";

                string DataFim = $"{Ano}1231";

                switch (Mes)
                {
                    case "13":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0101";

                            DataFim = $"{Ano}1231";
                            #endregion
                            break;
                        }
                    case "14":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0101";

                            DataFim = $"{Ano}0331";
                            #endregion
                            break;
                        }
                    case "15":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0401";

                            DataFim = $"{Ano}0631";
                            #endregion
                            break;
                        }
                    case "16":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0701";

                            DataFim = $"{Ano}0931";
                            #endregion
                            break;
                        }
                    case "17":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}1001";

                            DataFim = $"{Ano}1231";
                            #endregion
                            break;
                        }
                    case "18":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0101";

                            DataFim = $"{Ano}0631";
                            #endregion
                            break;
                        }
                    case "19":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0701";

                            DataFim = $"{Ano}1231";
                            #endregion
                            break;
                        }
                    default:
                        {
                            #region Parametros
                            DataInicio = $"{Ano}{Mes}01";

                            DataFim = $"{Ano}{Mes}31";

                            #endregion

                            break;
                        }
                }

                #region Faturados

                #region SubDistribuidorInter
                var FaturadoSubInter = (from SD10 in ProtheusInter.Sd2010s
                                        join SF10 in ProtheusInter.Sf4010s on SD10.D2Tes equals SF10.F4Codigo
                                        join SB10 in ProtheusInter.Sb1010s on SD10.D2Cod equals SB10.B1Cod
                                        join SC60 in ProtheusInter.Sc6010s on new { Filial = SD10.D2Filial, NumP = SD10.D2Pedido, Item = SD10.D2Itempv } equals new { Filial = SC60.C6Filial, NumP = SC60.C6Num, Item = SC60.C6Item }
                                        join SC50 in ProtheusInter.Sc5010s on new { Filial2 = SC60.C6Filial, NumC = SC60.C6Num } equals new { Filial2 = SC50.C5Filial, NumC = SC50.C5Num }
                                        join SA10 in ProtheusInter.Sa1010s on new { Codigo = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                                        join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                        where SD10.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*"
                                        && SF10.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*"
                                        && SF10.F4Duplic == "S" && SA10.A1Clinter == "S"
                                        && (int)(object)SD10.D2Emissao >= (int)(object)DataInicio && (int)(object)SD10.D2Emissao <= (int)(object)DataFim
                                        orderby SD10.D2Emissao
                                        select new RelatorioSubDistribuidor
                                        {
                                            Nome = SA10.A1Nreduz,
                                            Total = SD10.D2Total,
                                            Descon = SD10.D2Descon,
                                            Fabricante = SB10.B1Fabric
                                        }).ToList();
                var dolar = ProtheusInter.Sm2010s.Where(x => x.DELET != "*" && x.M2Moeda2 != 0).OrderByDescending(x => x.M2Data).FirstOrDefault();
                Faturamento.Add(new ValoresEmAberto { Nome = "INTERMEDIC", Valor = FaturadoSubInter.Sum(c => c.Total) });
                #endregion

                #region SubDistribuidorDenuo

                var FaturadoSubDenuo = (from SD10 in ProtheusDenuo.Sd2010s
                                        join SF10 in ProtheusDenuo.Sf4010s on SD10.D2Tes equals SF10.F4Codigo
                                        join SB10 in ProtheusDenuo.Sb1010s on SD10.D2Cod equals SB10.B1Cod
                                        join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SD10.D2Filial, NumP = SD10.D2Pedido, Item = SD10.D2Itempv } equals new { Filial = SC60.C6Filial, NumP = SC60.C6Num, Item = SC60.C6Item }
                                        join SC50 in ProtheusDenuo.Sc5010s on new { Filial2 = SC60.C6Filial, NumC = SC60.C6Num } equals new { Filial2 = SC50.C5Filial, NumC = SC50.C5Num }
                                        join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                                        join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                        where SD10.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*"
                                        && SF10.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SB10.DELET != "*"
                                        && SF10.F4Duplic == "S" && SA10.A1Clinter == "S"
                                        && (int)(object)SD10.D2Emissao >= (int)(object)DataInicio && (int)(object)SD10.D2Emissao <= (int)(object)DataFim
                                        orderby SD10.D2Emissao
                                        select new RelatorioSubDistribuidor
                                        {
                                            Nome = SA10.A1Nreduz,
                                            Total = SD10.D2Total,
                                            Descon = SD10.D2Descon,
                                            Fabricante = SB10.B1Fabric
                                        }).ToList();

                Faturamento.Add(new ValoresEmAberto { Nome = "DENUO", Valor = FaturadoSubDenuo.Sum(c => c.Total) });


                var cliente = FaturadoSubDenuo.GroupBy(x => x.Nome).Select(x => new
                {
                    Nome = x.Key,
                    Total = x.Sum(c => c.Total)
                }).OrderByDescending(x=> x.Total).ToList();

                var clienteInter = FaturadoSubInter.GroupBy(x => x.Nome).Select(x => new
                {
                    Nome = x.Key,
                    Total = x.Sum(c => c.Total)
                }).OrderByDescending(x => x.Total).ToList();

                var fabricanteInter = FaturadoSubInter.GroupBy(x => x.Fabricante).Select(x => new
                {
                    Nome = x.Key,
                    Total = x.Sum(c => c.Total)
                }).OrderByDescending(x => x.Total).ToList();

                var fabricanteDenuo = FaturadoSubDenuo.GroupBy(x => x.Fabricante).Select(x => new
                {
                    Nome = x.Key,
                    Total = x.Sum(c => c.Total)
                }).OrderByDescending(x => x.Total).ToList();

                cliente.ForEach(x =>
                {
                    ClientesDenuo.Add(new ValoresEmAberto { Nome = x.Nome, Valor = x.Total });
                });

                clienteInter.ForEach(x =>
                {
                   ClientesInter.Add(new ValoresEmAberto { Nome = x.Nome, Valor = x.Total });
                });

                fabricanteDenuo.ForEach(x =>
                {
                    FabricanteDenuo.Add(new ValoresEmAberto { Nome = x.Nome, Valor = x.Total });
                });

                fabricanteInter.ForEach(x =>
                {
                    FabricanteInter.Add(new ValoresEmAberto { Nome = x.Nome, Valor = x.Total });
                });

                #endregion

                #endregion


                var valores = new
                {
                    Valores = Faturamento,
                    ClientesDenuo = ClientesDenuo.OrderByDescending(x => x.Valor),
                    ClientesInter = ClientesInter.OrderByDescending(x => x.Valor),
                    FabricanteDenuo = FabricanteDenuo.OrderByDescending(x => x.Valor),
                    FabricanteInter = FabricanteInter.OrderByDescending(x => x.Valor),
                    ValorTotal = Faturamento.Sum(x => x.Valor)
                };


                var Teste = new
                {
                    valores = valores,
                };

                return new JsonResult(Teste);

            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex, SGID, "DashBoardMetas Faturados", user);

                return new JsonResult("");
            }
        }

        public JsonResult OnPostEmAberto(string Mes, string Ano)
        {

            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                string[] CF = new string[] { "5551", "6551", "6107", "6109" };
                string DataInicio = "";
                string DataFim = "";

                switch (Mes)
                {
                    case "13":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0101";

                            DataFim = $"{Ano}1231";
                            #endregion
                            break;
                        }
                    case "14":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0101";

                            DataFim = $"{Ano}0331";
                            #endregion
                            break;
                        }
                    case "15":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0401";

                            DataFim = $"{Ano}0631";
                            #endregion
                            break;
                        }
                    case "16":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0701";

                            DataFim = $"{Ano}0931";
                            #endregion
                            break;
                        }
                    case "17":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}1001";

                            DataFim = $"{Ano}1231";
                            #endregion
                            break;
                        }
                    case "18":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0101";

                            DataFim = $"{Ano}0631";
                            #endregion
                            break;
                        }
                    case "19":
                        {
                            #region Parametros

                            DataInicio = $"{Ano}0701";

                            DataFim = $"{Ano}1231";
                            #endregion
                            break;
                        }
                    default:
                        {
                            #region Parametros
                            DataInicio = $"{Ano}{Mes}01";

                            DataFim = $"{Ano}{Mes}31";

                            #endregion

                            break;
                        }
                }


                #region EmAberto

                #region Sub Inter
                var SubdistribuidorInter = (from SC50 in ProtheusInter.Sc5010s
                                            join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                            join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                            join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                            join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                            where SC50.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC60.DELET != "*" &&
                                            SA30.DELET != "*" && SB10.DELET != "*" && SA10.A1Clinter == "S" && SC50.C5Nota == "" &&
                                            SC60.C6Qtdven - SC60.C6Qtdent != 0
                                            && SA10.A1Cgc.Substring(0, 8) != "04715053"
                                            orderby SA10.A1Nome, SC50.C5Emissao
                                            select (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven
                             ).Sum();

                var dolar = ProtheusInter.Sm2010s.Where(x => x.DELET != "*" && x.M2Moeda2 != 0).OrderByDescending(x => x.M2Data).FirstOrDefault();

                ValoresEmAberto.Add(new ValoresEmAberto { Nome = "INTERMEDIC", Valor = SubdistribuidorInter * dolar.M2Moeda2 });
                #endregion

                #region Sub Denuo
                var SubdistribuidorDenuo = (from SC50 in ProtheusDenuo.Sc5010s
                                            join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                            join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                            join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                                            join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                                            where SC50.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC60.DELET != "*" &&
                                            SA30.DELET != "*" && SB10.DELET != "*" && SA10.A1Clinter == "S" && SC50.C5Nota == "" &&
                                            SC60.C6Qtdven - SC60.C6Qtdent != 0
                                            orderby SA10.A1Nome, SC50.C5Emissao
                                            select (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven
                             ).Sum();

                var dolar2 = ProtheusDenuo.Sm2010s.Where(x => x.DELET != "*" && x.M2Moeda2 != 0).OrderByDescending(x => x.M2Data).FirstOrDefault();

                ValoresEmAberto.Add(new ValoresEmAberto { Nome = "DENUO", Valor = SubdistribuidorDenuo * dolar2.M2Moeda2 });
                #endregion

                #endregion


                var valores = new
                {
                    Valores = ValoresEmAberto,
                    ValorTotal = ValoresEmAberto.Sum(x => x.Valor)
                };

                return new JsonResult(valores);
            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex, SGID, "DashBoardMetas EmAberto", user);

                return new JsonResult("");
            }

        }
    }
}
