using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Inter;
using SGID.Models.SubDistribuidor;

namespace SGID.Pages.Relatorios.SubDistribuidor
{
    [Authorize]
    public class SubDistribuidorPorFabricanteInterModel : PageModel
    {
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public string Fabricante { get; set; }
        public List<string> Fabricantes { get; set; }
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext Protheus { get; set; }
        public List<RelatorioSubDistribuidorPorFabricante> Relatorio { get; set; } = new List<RelatorioSubDistribuidorPorFabricante>();

        public SubDistribuidorPorFabricanteInterModel(TOTVSINTERContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public void OnGet()
        {
            Fabricantes = Protheus.Sa2010s.Where(x => x.DELET != "*").Select(x => x.A2Nreduz).ToList();
        }

        public IActionResult OnPostAsync(string Fabricante, DateTime DataInicio, DateTime DataFim)
        {
            this.Fabricante = Fabricante;
            Inicio = DataInicio;
            Fim = DataFim;

            Fabricantes = Protheus.Sa2010s.Where(x => x.DELET != "*").Select(x => x.A2Nreduz).ToList();

            Relatorio = Protheus.Sf2010s.Join(Protheus.Sd2010s, p => p.F2Doc, c => c.D2Doc, (p, c) => new { p, c })
                .Join(Protheus.Sb1010s, v => v.c.D2Cod, d => d.B1Cod, (q, d) => new { q, d })
                .Join(Protheus.Sa1010s, v => v.q.c.D2Cliente, c => c.A1Cod, (v, c) => new { v, c })
                .Where(x => x.v.q.p.DELET != "*" && x.v.q.p.F2Vend1 == "000065" && x.v.d.B1Fabric == this.Fabricante
                 && (int)(object)x.v.q.p.F2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)x.v.q.p.F2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", ""))
                .Select(x => new RelatorioSubDistribuidorPorFabricante
                {
                    Codigo = x.c.A1Cod,
                    Nome = x.c.A1Nome,
                    NF = x.v.q.p.F2Doc,
                    Emissao = x.v.q.p.F2Emissao,
                    Fabricante = x.v.d.B1Fabric,
                    CodProduto = x.v.q.c.D2Cod,
                    Descricao = x.v.d.B1Desc,
                    Quant = x.v.q.c.D2Quant,
                    PrcVen = x.v.q.c.D2Prcven,
                    Total = x.v.q.c.D2Total,
                    NomeVendedor = "PLINIO FERRAZ ALVIM"
                }).ToList();

            return Page();
        }

        public IActionResult OnPostExport(string Fabricante, DateTime DataInicio, DateTime DataFim)
        {
            try
            {
                this.Fabricante = Fabricante;
                Inicio = DataInicio;
                Fim = DataFim;

                Fabricantes = Protheus.Sa2010s.Where(x => x.DELET != "*").Select(x => x.A2Nreduz).ToList();

                Relatorio = Protheus.Sf2010s.Join(Protheus.Sd2010s, p => p.F2Doc, c => c.D2Doc, (p, c) => new { p, c })
                    .Join(Protheus.Sb1010s, v => v.c.D2Cod, d => d.B1Cod, (q, d) => new { q, d })
                    .Join(Protheus.Sa1010s, v => v.q.c.D2Cliente, c => c.A1Cod, (v, c) => new { v, c })
                    .Where(x => x.v.q.p.DELET != "*" && x.v.q.p.F2Vend1 == "000065" && x.v.d.B1Fabric == this.Fabricante
                     && (int)(object)x.v.q.p.F2Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "") && (int)(object)x.v.q.p.F2Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", ""))
                    .Select(x => new RelatorioSubDistribuidorPorFabricante
                    {
                        Codigo = x.c.A1Cod,
                        Nome = x.c.A1Nome,
                        NF = x.v.q.p.F2Doc,
                        Emissao = x.v.q.p.F2Emissao,
                        Fabricante = x.v.d.B1Fabric,
                        CodProduto = x.v.q.c.D2Cod,
                        Descricao = x.v.d.B1Desc,
                        Quant = x.v.q.c.D2Quant,
                        PrcVen = x.v.q.c.D2Prcven,
                        Total = x.v.q.c.D2Total,
                        NomeVendedor = "PLINIO FERRAZ ALVIM"
                    }).ToList();

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("SubDistribuidor Por Fabricante");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "SubDistribuidor Por Fabricante");

                sheet.Cells[1, 1].Value = "Cod.";
                sheet.Cells[1, 2].Value = "Nome";
                sheet.Cells[1, 3].Value = "NF";
                sheet.Cells[1, 4].Value = "Emissao";
                sheet.Cells[1, 5].Value = "Fabricante";
                sheet.Cells[1, 6].Value = "Cod.Produto";
                sheet.Cells[1, 7].Value = "Quant.";
                sheet.Cells[1, 8].Value = "Prc.Ven";
                sheet.Cells[1, 9].Value = "Total";
                sheet.Cells[1, 10].Value = "Vendedor";

                int i = 2;

                Relatorio.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.Codigo;
                    sheet.Cells[i, 2].Value = Pedido.Nome;
                    sheet.Cells[i, 3].Value = Pedido.NF;
                    sheet.Cells[i, 4].Value = Pedido.Emissao;
                    sheet.Cells[i, 5].Value = Pedido.Fabricante;
                    sheet.Cells[i, 6].Value = Pedido.CodProduto;
                    sheet.Cells[i, 7].Value = Pedido.Quant;
                    sheet.Cells[i, 8].Value = Pedido.PrcVen;
                    sheet.Cells[i, 9].Value = Pedido.Total;
                    sheet.Cells[i, 10].Value = Pedido.NomeVendedor;

                    i++;
                });

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "SubDistribuidorPorFabricanteInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SubDistribuidorPorFabricanteInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
