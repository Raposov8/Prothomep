using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;
using SGID.Models.Controladoria;
using SGID.Models.Denuo;
using System;

namespace SGID.Pages.DashBoards
{
    public class DashBoardSubDistribuidorModel : PageModel
    {
        public int FaturadoMes { get; set; }
        public double FaturadoMesValor { get; set; }
        public int EmAberto { get; set; }
        public double EmAbertoValor { get; set; }
        public int Baixado { get; set; }
        public double ValorBaixado { get; set; }
        public double Meta { get; set; }
        public double Comissao { get; set; }

        public TOTVSDENUOContext ProtheusDenuo { get; set; }
        public TOTVSINTERContext ProtheusInter { get; set; }

        public ApplicationDbContext SGID { get; set; }
        public DashBoardSubDistribuidorModel(TOTVSDENUOContext denuo, TOTVSINTERContext inter, ApplicationDbContext sgid)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
            SGID = sgid;
        }

        public void OnGet()
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                string data = DateTime.Now.ToString("yyyy/MM").Replace("/", "");
                string DataInicio = data + "01";
                string DataFim =  data + "31";
                int[] CF = new int[] { 5551, 6551, 6107, 6109 };

                var time = SGID.Times.FirstOrDefault(x => x.Integrante == user.ToLower());

                #region Intermedic

                var FaturadoInter = (from SD10 in ProtheusInter.Sd2010s
                             join SF10 in ProtheusInter.Sf4010s on SD10.D2Tes equals SF10.F4Codigo
                             join SB10 in ProtheusInter.Sb1010s on SD10.D2Cod equals SB10.B1Cod
                             join SC60 in ProtheusInter.Sc6010s on new { Filial = SD10.D2Filial, NumP = SD10.D2Pedido, Item = SD10.D2Itempv } equals new { Filial = SC60.C6Filial, NumP = SC60.C6Num, Item = SC60.C6Item }
                             join SC50 in ProtheusInter.Sc5010s on new { Filial2 = SC60.C6Filial, NumC = SC60.C6Num } equals new { Filial2 = SC50.C5Filial, NumC = SC50.C5Num }
                             join SA10 in ProtheusInter.Sa1010s on new { Codigo = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SD10.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*"
                             && SF10.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*"
                             && SF10.F4Duplic == "S" && SA10.A1Clinter == "S"
                             && (int)(object)SD10.D2Emissao >= (int)(object)DataInicio
                             && (int)(object)SD10.D2Emissao <= (int)(object)DataFim
                             orderby SD10.D2Emissao
                             select new RelatorioSubDistribuidor
                             {
                                 Nome = SA30.A3Nome,
                                 Nreduz = SA10.A1Nreduz,
                                 Doc = SD10.D2Doc,
                                 Cod = SB10.B1Cod,
                                 Desc = SB10.B1Desc,
                                 Quant = SD10.D2Quant,
                                 Total = SD10.D2Total,
                                 Descon = SD10.D2Descon,
                                 Emissao = $"{SD10.D2Emissao.Substring(6, 2)}/{SD10.D2Emissao.Substring(4, 2)}/{SD10.D2Emissao.Substring(0, 4)}",
                                 Num = SC50.C5Num,
                                 Utpoper = SC50.C5Utpoper,
                                 Fabricante = SB10.B1Fabric
                             }
                            ).ToList();

                var EmAbertoInter = (from SC50 in ProtheusInter.Sc5010s
                             join SA10 in ProtheusInter.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SC60 in ProtheusInter.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                             join SB10 in ProtheusInter.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                             join SA30 in ProtheusInter.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SC50.DELET != "*" && SA10.DELET != "*" && SC60.DELET != "*" &&
                             SA30.DELET != "*" && SA10.A1Clinter == "S" && SC50.C5Nota == "" &&
                             SC60.C6Qtdven - SC60.C6Qtdent != 0
                             && SA10.A1Cgc.Substring(0, 8) != "04715053"
                             orderby SA10.A1Nome, SC50.C5Emissao
                             select new RelatorioSubDistribuidor
                             {
                                 Num = SC50.C5Num,
                                 Emissao = $"{SC50.C5Emissao.Substring(6, 2)}/{SC50.C5Emissao.Substring(4, 2)}/{SC50.C5Emissao.Substring(0, 4)}",
                                 Nome = SA30.A3Nome,
                                 Desc = SB10.B1Desc,
                                 Quant = SC60.C6Qtdven - SC60.C6Qtdent,
                                 Doc = SC60.C6Produto,
                                 Total = (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven,
                                 Descon = SC60.C6Valdesc,
                                 Nreduz = SA10.A1Nome,
                                 Utpoper = SC50.C5Liberok,
                                 Fabricante = SB10.B1Fabric
                             }
                            ).ToList();

                #endregion

                #region Denuo

                var FaturadoDenuo = (from SD10 in ProtheusDenuo.Sd2010s
                             join SF10 in ProtheusDenuo.Sf4010s on SD10.D2Tes equals SF10.F4Codigo
                             join SB10 in ProtheusDenuo.Sb1010s on SD10.D2Cod equals SB10.B1Cod
                             join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SD10.D2Filial, NumP = SD10.D2Pedido, Item = SD10.D2Itempv } equals new { Filial = SC60.C6Filial, NumP = SC60.C6Num, Item = SC60.C6Item }
                             join SC50 in ProtheusDenuo.Sc5010s on new { Filial2 = SC60.C6Filial, NumC = SC60.C6Num } equals new { Filial2 = SC50.C5Filial, NumC = SC50.C5Num }
                             join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SD10.DELET != "*" && SC60.DELET != "*" && SC50.DELET != "*"
                             && SF10.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*"
                             && SF10.F4Duplic == "S" && SA10.A1Clinter == "S"
                             && (int)(object)SD10.D2Emissao >= (int)(object)DataInicio 
                             && (int)(object)SD10.D2Emissao <= (int)(object)DataFim
                             orderby SD10.D2Emissao
                             select new RelatorioSubDistribuidor
                             {
                                 Nome = SA30.A3Nome,
                                 Nreduz = SA10.A1Nreduz,
                                 Doc = SD10.D2Doc,
                                 Cod = SB10.B1Cod,
                                 Desc = SB10.B1Desc,
                                 Quant = SD10.D2Quant,
                                 Total = SD10.D2Total,
                                 Descon = SD10.D2Descon,
                                 Emissao = $"{SD10.D2Emissao.Substring(6, 2)}/{SD10.D2Emissao.Substring(4, 2)}/{SD10.D2Emissao.Substring(0, 4)}",
                                 Num = SC50.C5Num,
                                 Utpoper = SC50.C5Utpoper,
                                 Fabricante = SB10.B1Fabric
                             }
                            ).ToList();

                var EmAbertoDenuo = (from SC50 in ProtheusDenuo.Sc5010s
                             join SA10 in ProtheusDenuo.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SC60 in ProtheusDenuo.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                             join SB10 in ProtheusDenuo.Sb1010s on SC60.C6Produto equals SB10.B1Cod
                             join SA30 in ProtheusDenuo.Sa3010s on SC50.C5Vend1 equals SA30.A3Cod
                             where SC50.DELET != "*" && SA10.DELET != "*" && SC60.DELET != "*" &&
                             SA30.DELET != "*" && SA10.A1Clinter == "S" && SC50.C5Nota == "" &&
                             SC60.C6Qtdven - SC60.C6Qtdent != 0
                             orderby SA10.A1Nome, SC50.C5Emissao
                             select new RelatorioSubDistribuidor
                             {
                                 Num = SC50.C5Num,
                                 Emissao = $"{SC50.C5Emissao.Substring(6, 2)}/{SC50.C5Emissao.Substring(4, 2)}/{SC50.C5Emissao.Substring(0, 4)}",
                                 Nome = SA30.A3Nome,
                                 Desc = SB10.B1Desc,
                                 Quant = SC60.C6Qtdven - SC60.C6Qtdent,
                                 Doc = SC60.C6Produto,
                                 Total = (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven,
                                 Descon = SC60.C6Valdesc,
                                 Nreduz = SA10.A1Nome,
                                 Utpoper = SC50.C5Liberok,
                                 Fabricante = SB10.B1Fabric
                             }
                            ).ToList();

                #endregion

                #region Baixa 

                var baixa = (from SE1 in ProtheusDenuo.Se1010s
                             join SA3 in ProtheusDenuo.Sa3010s on SE1.E1Vend1 equals SA3.A3Cod
                             where (int)(object)SE1.E1Baixa >= (int)(object)DataInicio
                             && (int)(object)SE1.E1Baixa <= (int)(object)DataFim
                             && SA3.A3Xlogin == "PLINIO.ALVIM"
                             select new
                             {
                                 SE1.E1Num,
                                 SE1.E1Parcela,
                                 SE1.E1Tipo,
                                 SE1.E1Naturez,
                                 SE1.E1Nomclie,
                                 SE1.E1Emissao,
                                 SE1.E1Valor,
                                 SE1.E1Baixa,
                                 SA3.A3Nome,
                                 SA3.A3Xlogin,
                                 SA3.A3Cod
                             }).OrderBy(x=>x.E1Num).ToList();

                var baixaInter = (from SE1 in ProtheusInter.Se1010s
                             join SA3 in ProtheusInter.Sa3010s on SE1.E1Vend1 equals SA3.A3Cod
                             where (int)(object)SE1.E1Baixa >= (int)(object)DataInicio
                             && (int)(object)SE1.E1Baixa <= (int)(object)DataFim
                             && SA3.A3Xlogin == "PLINIO.ALVIM"
                             select new
                             {
                                 SE1.E1Num,
                                 SE1.E1Parcela,
                                 SE1.E1Tipo,
                                 SE1.E1Naturez,
                                 SE1.E1Nomcli,
                                 SE1.E1Emissao,
                                 SE1.E1Valor,
                                 SE1.E1Baixa,
                                 SA3.A3Nome,
                                 SA3.A3Xlogin,
                                 SA3.A3Cod
                             }).OrderBy(x => x.E1Num).ToList();

                #endregion

                FaturadoMes = FaturadoInter.DistinctBy(x=>x.Num).Count() + FaturadoDenuo.DistinctBy(x => x.Num).Count();
                FaturadoMesValor = FaturadoInter.Sum(x => x.Total) + FaturadoDenuo.Sum(x => x.Total);

                EmAberto = EmAbertoInter.DistinctBy(x => x.Num).Count() + EmAbertoDenuo.DistinctBy(x => x.Num).Count();
                EmAbertoValor = EmAbertoInter.Sum(x => x.Total) + EmAbertoDenuo.Sum(x => x.Total);

                Baixado = baixa.DistinctBy(x => x.E1Num).Count() + baixaInter.DistinctBy(x => x.E1Num).Count();
                ValorBaixado = baixa.Sum(x => x.E1Valor) + baixaInter.Sum(x => x.E1Valor);

                Meta = time.Meta;
                Comissao = ValorBaixado * (time.Porcentagem / 100);

            }
            catch (Exception ex)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(ex, SGID, "DashBoardSubDistribuidor", user);
            }
        }
    }
}
