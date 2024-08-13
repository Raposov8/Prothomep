using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models;

namespace SGID.Pages.Relatorios.Cirurgias
{
    [Authorize]
    public class CirurgiasAFaturarInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioCirurgiasFaturar> Relatorio { get; set; } = new List<RelatorioCirurgiasFaturar>();

        public double Total { get; set; }

        public CirurgiasAFaturarInterModel(TOTVSINTERContext context,ApplicationDbContext sgid)
        {
            Protheus = context;
            SGID = sgid;
        }
        public void OnGet()
        {
            try
            {

                var user = User.Identity.Name.Split("@")[0].ToUpper();

                Relatorio = (from SC5 in Protheus.Sc5010s
                             from SC6 in Protheus.Sc6010s
                             from SA1 in Protheus.Sa1010s
                             from SA3 in Protheus.Sa3010s
                             from SF4 in Protheus.Sf4010s
                             from SB1 in Protheus.Sb1010s
                             where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                             && SC6.C6Nota == "" && SC5.C5Nota == "" && SC6.C6Blq != "R" && SC6.DELET != "*" && SF4.F4Filial == SC6.C6Filial
                             && SF4.F4Codigo == SC6.C6Tes && SF4.F4Duplic == "S" && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                             && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                             && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T") && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                             && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140" && SA3.A3Xlogin == user
                             orderby SC5.C5Num descending
                             select new RelatorioCirurgiasFaturar
                             {
                                 Filial = SC5.C5Filial,
                                 Num = SC5.C5Num,
                                 Unumage = SC5.C5Unumage,
                                 XNomMed = SC5.C5XNmmed,
                                 XDtCir = SC5.C5XDtcir,
                                 UPatrim = SC6.C6Upatrim,
                                 UProces = SC5.C5Uproces,
                                 XNomPac = SC5.C5XNmpac,
                                 UTPoper = SC5.C5Utpoper,
                                 NomeCliente = SA1.A1Nome,
                                 TipoOper = SC5.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC5.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : SC5.C5Utpoper == "T" ? "TRIAGEM DE PEDIDOS" : "OUTROS",
                                 Produto = SC6.C6Produto,
                                 Desc = SB1.B1Desc,
                                 NomeVendedor = SC5.C5Nomvend,
                                 QtdVen = SC6.C6Qtdven,
                                 PrcVen = SC6.C6Prcven,
                                 QtdDent = SC6.C6Qtdent,
                                 Valor = SC6.C6Valor,
                                 Login = SA3.A3Xlogin,
                                 Faturado = SC5.C5Nota != "" || SC5.C5Liberok == "E" && SC5.C5Blq == "" ? "S" : "N",
                                 Emissao = $"{SC5.C5Emissao.Substring(4, 2)}/{SC5.C5Emissao.Substring(4, 2)}/{SC5.C5Emissao.Substring(0, 4)}",
                                 XDtAut = SC5.C5XDtaut
                             }
                            ).ToList();

                Relatorio.ForEach(x =>
                {
                    Total += x.Valor;
                });
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasAFaturarInter",user);
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {
                var user = User.Identity.Name.Split("@")[0].ToUpper();

                Relatorio = (from SC5 in Protheus.Sc5010s
                             from SC6 in Protheus.Sc6010s
                             from SA1 in Protheus.Sa1010s
                             from SA3 in Protheus.Sa3010s
                             from SF4 in Protheus.Sf4010s
                             from SB1 in Protheus.Sb1010s
                             where SC5.DELET != "*" && SC6.C6Filial == SC5.C5Filial && SC6.C6Num == SC5.C5Num
                             && SC6.C6Nota == "" && SC5.C5Nota == "" && SC6.C6Blq != "R" && SC6.DELET != "*" && SF4.F4Filial == SC6.C6Filial
                             && SF4.F4Codigo == SC6.C6Tes && SF4.F4Duplic == "S" && SF4.DELET != "*" && SA1.A1Cod == SC5.C5Cliente
                             && SA1.A1Loja == SC5.C5Lojacli && SA1.DELET != "*" && SA3.A3Cod == SC5.C5Vend1 && SA3.DELET != "*"
                             && (SC5.C5Utpoper == "F" || SC5.C5Utpoper == "T") && SB1.DELET != "*" && SC6.C6Produto == SB1.B1Cod
                             && SA1.A1Cgc != "04715053000140" && SA1.A1Cgc != "04715053000220" && SA1.A1Cgc != "01390500000140" && SA3.A3Xlogin == user
                             orderby SC5.C5Num descending
                             select new RelatorioCirurgiasFaturar
                             {
                                 Filial = SC5.C5Filial,
                                 Num = SC5.C5Num,
                                 Unumage = SC5.C5Unumage,
                                 XNomMed = SC5.C5XNmmed,
                                 XDtCir = SC5.C5XDtcir,
                                 UPatrim = SC6.C6Upatrim,
                                 UProces = SC5.C5Uproces,
                                 XNomPac = SC5.C5XNmpac,
                                 UTPoper = SC5.C5Utpoper,
                                 NomeCliente = SA1.A1Nome,
                                 TipoOper = SC5.C5Utpoper == "F" ? "FATURAMENTO DE CIRURGIAS" : SC5.C5Utpoper == "V" ? "VENDA PARA SUB-DISTRIBUIDOR" : SC5.C5Utpoper == "T" ? "TRIAGEM DE PEDIDOS" : "OUTROS",
                                 Produto = SC6.C6Produto,
                                 Desc = SB1.B1Desc,
                                 NomeVendedor = SC5.C5Nomvend,
                                 QtdVen = SC6.C6Qtdven,
                                 PrcVen = SC6.C6Prcven,
                                 QtdDent = SC6.C6Qtdent,
                                 Valor = SC6.C6Valor,
                                 Login = SA3.A3Xlogin,
                                 Faturado = SC5.C5Nota != "" || SC5.C5Liberok == "E" && SC5.C5Blq == "" ? "S" : "N",
                                 Emissao = $"{SC5.C5Emissao.Substring(4, 2)}/{SC5.C5Emissao.Substring(4, 2)}/{SC5.C5Emissao.Substring(0, 4)}",
                                 XDtAut = SC5.C5XDtaut
                             }
                            ).ToList();

                Relatorio.ForEach(x =>
                {
                    Total += x.Valor;
                });

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("Cirurgias A Faturar");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Cirurgias A Faturar");

                sheet.Cells[1, 1].Value = "Vendedor";
                sheet.Cells[1, 2].Value = "Num.Pedido";
                sheet.Cells[1, 3].Value = "Data Cirurgia";
                sheet.Cells[1, 4].Value = "Filial";
                sheet.Cells[1, 5].Value = "Emissão";
                sheet.Cells[1, 6].Value = "Cliente";
                sheet.Cells[1, 7].Value = "Agendamento";
                sheet.Cells[1, 8].Value = "Tipo";
                sheet.Cells[1, 9].Value = "Médico";
                sheet.Cells[1, 10].Value = "Paciente";
                sheet.Cells[1, 11].Value = "Produto";
                sheet.Cells[1, 12].Value = "Desc.Produto";
                sheet.Cells[1, 13].Value = "Qtd.Pedido";
                sheet.Cells[1, 14].Value = "Vl.Unitario";
                sheet.Cells[1, 15].Value = "Total Pedido";
                sheet.Cells[1, 16].Value = "DT.Autorização";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.NomeVendedor;
                    sheet.Cells[i, 2].Value = Pedido.Num;
                    sheet.Cells[i, 3].Value = Pedido.XDtCir;
                    sheet.Cells[i, 4].Value = Pedido.Filial;
                    sheet.Cells[i, 5].Value = Pedido.Emissao;
                    sheet.Cells[i, 6].Value = Pedido.NomeCliente;
                    sheet.Cells[i, 7].Value = Pedido.Unumage;
                    sheet.Cells[i, 8].Value = Pedido.TipoOper;
                    sheet.Cells[i, 9].Value = Pedido.XNomMed;
                    sheet.Cells[i, 10].Value = Pedido.XNomPac;
                    sheet.Cells[i, 11].Value = Pedido.Produto;
                    sheet.Cells[i, 12].Value = Pedido.Desc;
                    sheet.Cells[i, 13].Value = Pedido.QtdVen;
                    sheet.Cells[i, 14].Value = Pedido.PrcVen;
                    sheet.Cells[i, 15].Value = Pedido.Valor;
                    sheet.Cells[i, 16].Value = Pedido.XDtAut;
                    i++;
                });


                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CirurgiaAFaturar.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiaAFaturar Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
