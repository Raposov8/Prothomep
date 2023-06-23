using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Comercial;
using SGID.Models.Denuo;

namespace SGID.Pages.Relatorios.Comercial
{
    [Authorize]
    public class NFSModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioNFS> Relatorios { get; set; } = new List<RelatorioNFS>();
        public NFSModel(TOTVSDENUOContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public int TotalPages { get; set; }
        public int CurrentPage { get; set; }
        
        public string CodCli { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPostAsync(string Cliente,int Pagina = 1)
        {
            try
            {
                CodCli = Cliente;
                CurrentPage = Pagina;

                var queryselect = (from SD20 in Protheus.Sd2010s
                                   join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                   where SD20.DELET != "*" && SB10.DELET != "*" && SD20.D2Cliente == Cliente
                                   orderby SD20.D2Emissao descending
                                   select new RelatorioNFS
                                   {
                                       Filial = SD20.D2Filial,
                                       NF = SD20.D2Doc,
                                       Serie = SD20.D2Serie,
                                       DTEmissao = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                                       CodCliente = SD20.D2Cliente,
                                       Loja = SD20.D2Loja,
                                       CodProd = SD20.D2Cod,
                                       Produto = SB10.B1Desc,
                                       QTDE = SD20.D2Quant,
                                       ValUnit = SD20.D2Prcven,
                                       Total = SD20.D2Total,
                                       CFOP = SD20.D2Cf

                                   });

                TotalPages = (int)Math.Ceiling(decimal.Divide(queryselect.Count(), 20));

                Relatorios = queryselect.Skip((Pagina - 1) * 20).Take(20).ToList();

                return Page();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NFS",user);

                return LocalRedirect("/error");
            }
        }
        public IActionResult OnPostExport()
        {
            try
            {

                var user = User.Identity.Name.Split("@")[0].ToUpper();
                var queryselect = (from SD20 in Protheus.Sd2010s
                                   join SB10 in Protheus.Sb1010s on SD20.D2Cod equals SB10.B1Cod
                                   where SD20.DELET != "*" && SB10.DELET != "*" && SD20.D2Cliente == CodCli
                                   orderby SD20.D2Emissao descending
                                   select new RelatorioNFS
                                   {
                                       Filial = SD20.D2Filial,
                                       NF = SD20.D2Doc,
                                       Serie = SD20.D2Serie,
                                       DTEmissao = $"{SD20.D2Emissao.Substring(6, 2)}/{SD20.D2Emissao.Substring(4, 2)}/{SD20.D2Emissao.Substring(0, 4)}",
                                       CodCliente = SD20.D2Cliente,
                                       Loja = SD20.D2Loja,
                                       CodProd = SD20.D2Cod,
                                       Produto = SB10.B1Desc,
                                       QTDE = SD20.D2Quant,
                                       ValUnit = SD20.D2Prcven,
                                       Total = SD20.D2Total,
                                       CFOP = SD20.D2Cf

                                   });

                Relatorios = queryselect.ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("NFS");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "NFS");

                sheet.Cells[1, 1].Value = "Filial";
                sheet.Cells[1, 2].Value = "NF";
                sheet.Cells[1, 3].Value = "Serie";
                sheet.Cells[1, 4].Value = "DT Emissão";
                sheet.Cells[1, 5].Value = "Cod.Cliente";
                sheet.Cells[1, 6].Value = "Loja";
                sheet.Cells[1, 7].Value = "Cod.Prod.";
                sheet.Cells[1, 8].Value = "Produto";
                sheet.Cells[1, 9].Value = "QTDE";
                sheet.Cells[1, 10].Value = "Val.Unit.";
                sheet.Cells[1, 11].Value = "Total";
                sheet.Cells[1, 12].Value = "CFOP";

                int i = 2;

                Relatorios.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Filial;
                    sheet.Cells[i, 2].Value = Pedido.NF;
                    sheet.Cells[i, 3].Value = Pedido.Serie;
                    sheet.Cells[i, 4].Value = Pedido.DTEmissao;
                    sheet.Cells[i, 5].Value = Pedido.CodCliente;
                    sheet.Cells[i, 6].Value = Pedido.Loja;
                    sheet.Cells[i, 7].Value = Pedido.CodProd;
                    sheet.Cells[i, 8].Value = Pedido.Produto;
                    sheet.Cells[i, 9].Value = Pedido.QTDE;
                    sheet.Cells[i, 10].Value = Pedido.ValUnit;
                    sheet.Cells[i, 11].Value = Pedido.Total;
                    sheet.Cells[i, 12].Value = Pedido.CFOP;

                    i++;
                });


                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "NFS.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "NFS Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
