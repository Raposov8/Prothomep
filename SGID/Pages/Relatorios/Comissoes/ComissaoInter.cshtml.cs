using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Comissoes;

namespace SGID.Pages.Relatorios.Comissoes
{
    [Authorize]
    public class ComissaoInterModel : PageModel
    {

        public TOTVSINTERContext Protheus { get; set; }
        public ApplicationDbContext SGID { get; set; }

        public List<RelatorioComissao> Relatorio { get; set; } = new List<RelatorioComissao>();

        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }
        public ComissaoInterModel(TOTVSINTERContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public void OnGet()
        {
            
        }

        public IActionResult OnPost(DateTime DataInicio,DateTime DataFim)
        {
            Inicio = DataInicio;
            Fim = DataFim;

            var cfs = new string[] { "5551", "6551", "6107", "6109" };

            var teste = (from SD20 in Protheus.Sd2010s
                         join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                         join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                         join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                         join SX50GRP in Protheus.Sx5010s on new { Chave = SA10.A1Xgrinte, Tabela = "Z3" } equals new { Chave = SX50GRP.X5Chave, Tabela = SX50GRP.X5Tabela }
                         join SX50UNN in Protheus.Sx5010s on new { Chave = SC50.C5Xunnego, Tabela = "Z8" } equals new { Chave = SX50UNN.X5Chave, Tabela = SX50UNN.X5Tabela }
                         where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SC50.DELET != "*" && SX50GRP.DELET != "*" && SX50UNN.DELET != "*" && 
                         (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/","") && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                         && SD20.D2Quant > 0 && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114
                         || (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || cfs.Contains(SD20.D2Cf))
                         select new RelatorioComissao
                         {
                             Filial = SD20.D2Filial,
                             Pedido = SD20.D2Pedido,
                             Cliente = SD20.D2Cliente,
                             Loja = SD20.D2Loja,
                             ItemPv = SD20.D2Itempv,
                             Cod = SD20.D2Cod,
                             XGrinte = SA10.A1Xgrinte,
                             GrupoCli = SX50GRP.X5Descri ?? SA10.A1Nreduz,
                             TipoCliInt = SA10.A1Clinter == "H" ? "HOSPITAL" : SA10.A1Clinter == "M" ? "MEDICO" : SA10.A1Clinter == "I" ? "INSTRUMENTADOR" : SA10.A1Clinter == "N" ? "NORMAL" : SA10.A1Clinter == "C" ? "CONVENIO" : SA10.A1Clinter == "P" ? "PARTICULAR" : SA10.A1Clinter == "S" ? "SUBDISTRIB" : SA10.A1Clinter == "G" ? "INTRA GRUPO" : SA10.A1Clinter == "D" ? "DENTAL" : "OUTROS",
                             Analista = SA10.A1Xatecli,
                             Empresa = "INTERMEDIC",
                             TotalFat = SD20.D2Total,
                             Valipi = SD20.D2Valipi,
                             Valicm = SD20.D2Valicm,
                             Descon = SD20.D2Descon,
                             DTEmisNF = SD20.D2Emissao,
                             AnoMesFat = SD20.D2Emissao,
                             AGPrinc = (from SC50b in Protheus.Sc5010s
                                        join PAC10 in Protheus.Pac010s on new { Filial = SC50b.C5Filial, UnuMage = SC50b.C5Unumage } equals new { Filial = PAC10.PacFilial, UnuMage = PAC10.PacNumage }
                                        where SC50b.DELET != "*" && PAC10.DELET != "*" && SC50b.C5Unumage != "" && SC50b.C5Filial == SD20.D2Filial && SC50b.C5Num == SD20.D2Pedido
                                        select PAC10.PacXcompl == "1" ? PAC10.PacXagepr : PAC10.PacNumage).First() ?? SD20.D2Filial + SD20.D2Pedido,
                             Fabric = SB10.B1Xfabric,
                             Xunnego = SC50.C5Xunnego,
                             Unnego = SX50UNN.X5Descri ?? "",
                             TipoPV = "O",
                             CF = SD20.D2Cf
                         }
                         );


            var teste1 = teste.GroupBy(x => new
            {
                x.Filial,
                x.Pedido,
                x.Cliente,
                x.Loja,
                x.ItemPv,
                x.Cod,
                x.TipoCliInt,
                x.XGrinte,
                x.GrupoCli,
                x.Analista,
                x.Fabric,
                x.Xunnego,
                x.Unnego,
                x.DTEmisNF,
                x.AnoMesFat,
                x.AGPrinc,
                x.TipoPV,
                x.Empresa,
                x.CF
            });

            Relatorio = teste1.Select(x => new RelatorioComissao
            {
                Filial = x.Key.Filial,
                Pedido = x.Key.Pedido,
                Cliente = x.Key.Cliente,
                Loja = x.Key.Loja,
                ItemPv = x.Key.ItemPv,
                Cod = x.Key.Cod,
                XGrinte = x.Key.XGrinte,
                GrupoCli = x.Key.GrupoCli,
                TipoCliInt = x.Key.TipoCliInt,
                Analista = x.Key.Analista,
                Empresa = x.Key.Empresa,
                TotalFat = x.Sum(c => c.TotalFat),
                Valipi = x.Sum(c => c.Valipi),
                Valicm = x.Sum(c => c.Valicm),
                Descon = x.Sum(c => c.Descon),
                DTEmisNF = $"{x.Key.DTEmisNF.Substring(6, 2)}/{x.Key.DTEmisNF.Substring(4, 2)}/{x.Key.DTEmisNF.Substring(0, 4)}",
                AnoMesFat = x.Key.AnoMesFat.Substring(0, 6),
                AGPrinc = x.Key.AGPrinc,
                Fabric = x.Key.Fabric,
                Xunnego = x.Key.Xunnego,
                Unnego = x.Key.Unnego,
                TipoPV = x.Key.TipoPV,
                CF = x.Key.CF
            }).ToList();

            return Page();
        }

        public IActionResult OnPostExport(DateTime DataInicio, DateTime DataFim)
        {
            try
            {

                Inicio = DataInicio;
                Fim = DataFim;

                var cfs = new string[] { "5551", "6551", "6107", "6109" };

                var teste = (from SD20 in Protheus.Sd2010s
                             join SA10 in Protheus.Sa1010s on new { Cod = SD20.D2Cliente, Loja = SD20.D2Loja } equals new { Cod = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                             join SC50 in Protheus.Sc5010s on new { Filial = SD20.D2Filial, Num = SD20.D2Pedido } equals new { Filial = SC50.C5Filial, Num = SC50.C5Num }
                             join SX50GRP in Protheus.Sx5010s on new { Chave = SA10.A1Xgrinte, Tabela = "Z3" } equals new { Chave = SX50GRP.X5Chave, Tabela = SX50GRP.X5Tabela }
                             join SX50UNN in Protheus.Sx5010s on new { Chave = SC50.C5Xunnego, Tabela = "Z8" } equals new { Chave = SX50UNN.X5Chave, Tabela = SX50UNN.X5Tabela }
                             where SD20.DELET != "*" && SA10.DELET != "*" && SB10.DELET != "*" && SC50.DELET != "*" && SX50GRP.DELET != "*" && SX50UNN.DELET != "*" &&
                             (int)(object)SD20.D2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)SD20.D2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                             && SD20.D2Quant > 0 && ((int)(object)SD20.D2Cf >= 5102 && (int)(object)SD20.D2Cf <= 5114 || (int)(object)SD20.D2Cf >= 6102 && (int)(object)SD20.D2Cf <= 6114
                             || (int)(object)SD20.D2Cf >= 7102 && (int)(object)SD20.D2Cf <= 7114 || cfs.Contains(SD20.D2Cf))
                             select new RelatorioComissao
                             {
                                 Filial = SD20.D2Filial,
                                 Pedido = SD20.D2Pedido,
                                 Cliente = SD20.D2Cliente,
                                 Loja = SD20.D2Loja,
                                 ItemPv = SD20.D2Itempv,
                                 Cod = SD20.D2Cod,
                                 XGrinte = SA10.A1Xgrinte,
                                 GrupoCli = SX50GRP.X5Descri ?? SA10.A1Nreduz,
                                 TipoCliInt = SA10.A1Clinter == "H" ? "HOSPITAL" : SA10.A1Clinter == "M" ? "MEDICO" : SA10.A1Clinter == "I" ? "INSTRUMENTADOR" : SA10.A1Clinter == "N" ? "NORMAL" : SA10.A1Clinter == "C" ? "CONVENIO" : SA10.A1Clinter == "P" ? "PARTICULAR" : SA10.A1Clinter == "S" ? "SUBDISTRIB" : SA10.A1Clinter == "G" ? "INTRA GRUPO" : SA10.A1Clinter == "D" ? "DENTAL" : "OUTROS",
                                 Analista = SA10.A1Xatecli,
                                 Empresa = "INTERMEDIC",
                                 TotalFat = SD20.D2Total,
                                 Valipi = SD20.D2Valipi,
                                 Valicm = SD20.D2Valicm,
                                 Descon = SD20.D2Descon,
                                 DTEmisNF = SD20.D2Emissao,
                                 AnoMesFat = SD20.D2Emissao,
                                 AGPrinc = (from SC50b in Protheus.Sc5010s
                                            join PAC10 in Protheus.Pac010s on new { Filial = SC50b.C5Filial, UnuMage = SC50b.C5Unumage } equals new { Filial = PAC10.PacFilial, UnuMage = PAC10.PacNumage }
                                            where SC50b.DELET != "*" && PAC10.DELET != "*" && SC50b.C5Unumage != "" && SC50b.C5Filial == SD20.D2Filial && SC50b.C5Num == SD20.D2Pedido
                                            select PAC10.PacXcompl == "1" ? PAC10.PacXagepr : PAC10.PacNumage).First() ?? SD20.D2Filial + SD20.D2Pedido,
                                 Fabric = SB10.B1Xfabric,
                                 Xunnego = SC50.C5Xunnego,
                                 Unnego = SX50UNN.X5Descri ?? "",
                                 TipoPV = "O",
                                 CF = SD20.D2Cf
                             }
                             );


                var teste1 = teste.GroupBy(x => new
                {
                    x.Filial,
                    x.Pedido,
                    x.Cliente,
                    x.Loja,
                    x.ItemPv,
                    x.Cod,
                    x.TipoCliInt,
                    x.XGrinte,
                    x.GrupoCli,
                    x.Analista,
                    x.Fabric,
                    x.Xunnego,
                    x.Unnego,
                    x.DTEmisNF,
                    x.AnoMesFat,
                    x.AGPrinc,
                    x.TipoPV,
                    x.Empresa,
                    x.CF
                });

                Relatorio = teste1.Select(x => new RelatorioComissao
                {
                    Filial = x.Key.Filial,
                    Pedido = x.Key.Pedido,
                    Cliente = x.Key.Cliente,
                    Loja = x.Key.Loja,
                    ItemPv = x.Key.ItemPv,
                    Cod = x.Key.Cod,
                    XGrinte = x.Key.XGrinte,
                    GrupoCli = x.Key.GrupoCli,
                    TipoCliInt = x.Key.TipoCliInt,
                    Analista = x.Key.Analista,
                    Empresa = x.Key.Empresa,
                    TotalFat = x.Sum(c => c.TotalFat),
                    Valipi = x.Sum(c => c.Valipi),
                    Valicm = x.Sum(c => c.Valicm),
                    Descon = x.Sum(c => c.Descon),
                    DTEmisNF = $"{x.Key.DTEmisNF.Substring(6, 2)}/{x.Key.DTEmisNF.Substring(4, 2)}/{x.Key.DTEmisNF.Substring(0, 4)}",
                    AnoMesFat = x.Key.AnoMesFat.Substring(0, 6),
                    AGPrinc = x.Key.AGPrinc,
                    Fabric = x.Key.Fabric,
                    Xunnego = x.Key.Xunnego,
                    Unnego = x.Key.Unnego,
                    TipoPV = x.Key.TipoPV,
                    CF = x.Key.CF
                }).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Comissao");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Comissao");

                sheet.Cells[1, 1].Value = "Filial";
                sheet.Cells[1, 2].Value = "Pedido";
                sheet.Cells[1, 3].Value = "Cliente";
                sheet.Cells[1, 4].Value = "Loja";
                sheet.Cells[1, 5].Value = "ItemPV";
                sheet.Cells[1, 6].Value = "Cod";
                sheet.Cells[1, 7].Value = "XGrinte";
                sheet.Cells[1, 8].Value = "Grupo Cli";
                sheet.Cells[1, 9].Value = "Tipo Cli Int";
                sheet.Cells[1, 10].Value = "Analista";
                sheet.Cells[1, 11].Value = "Empresa";
                sheet.Cells[1, 12].Value = "Total Fat";
                sheet.Cells[1, 13].Value = "Valipi";
                sheet.Cells[1, 14].Value = "Valicm";
                sheet.Cells[1, 15].Value = "Descon";
                sheet.Cells[1, 16].Value = "DTEmisNF";
                sheet.Cells[1, 17].Value = "AnoMesFat";
                sheet.Cells[1, 18].Value = "AGPrinc";
                sheet.Cells[1, 19].Value = "Fabric";
                sheet.Cells[1, 20].Value = "XUNNEGO";
                sheet.Cells[1, 21].Value = "UNNEGO";
                sheet.Cells[1, 22].Value = "TipoPV";
                sheet.Cells[1, 23].Value = "CF";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.Pedido;
                    sheet.Cells[i, 3].Value = Pedido.Cliente;
                    sheet.Cells[i, 4].Value = Pedido.Loja;
                    sheet.Cells[i, 5].Value = Pedido.ItemPv;
                    sheet.Cells[i, 6].Value = Pedido.Cod;
                    sheet.Cells[i, 7].Value = Pedido.XGrinte;
                    sheet.Cells[i, 8].Value = Pedido.GrupoCli;
                    sheet.Cells[i, 9].Value = Pedido.TipoCliInt;
                    sheet.Cells[i, 10].Value = Pedido.Analista;
                    sheet.Cells[i, 11].Value = Pedido.Empresa;
                    sheet.Cells[i, 12].Value = Pedido.TotalFat;
                    sheet.Cells[i, 13].Value = Pedido.Valipi;
                    sheet.Cells[i, 14].Value = Pedido.Valicm;
                    sheet.Cells[i, 15].Value = Pedido.Descon;
                    sheet.Cells[i, 16].Value = Pedido.DTEmisNF;
                    sheet.Cells[i, 17].Value = Pedido.AnoMesFat;
                    sheet.Cells[i, 18].Value = Pedido.AGPrinc;
                    sheet.Cells[i, 19].Value = Pedido.Fabric;
                    sheet.Cells[i, 20].Value = Pedido.Xunnego;
                    sheet.Cells[i, 21].Value = Pedido.Unnego;
                    sheet.Cells[i, 22].Value = Pedido.TipoPV;
                    sheet.Cells[i, 23].Value = Pedido.CF;

                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ComissaoInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ComissaoInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
