using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.DTO;
using SGID.Models.Denuo;
using SGID.Models.Services;
using Microsoft.AspNetCore.Authorization;
using SGID.Models.Inter;

namespace SGID.Pages.Cirurgias
{
    [Authorize]
    public class PdfAgendamentoModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private readonly IWebHostEnvironment _WEB;
        public Agendamentos Agendamento { get; set; }

        public List<Produto> Produtos { get; set; } = new List<Produto>();
        public double UnidadesTotal { get; set; }
        public double ValorTotal { get; set; }

        public PdfAgendamentoModel(ApplicationDbContext sGID,TOTVSDENUOContext denuo, TOTVSINTERContext inter, IWebHostEnvironment wEB)
        {
            SGID = sGID;
            ProtheusInter = inter;
            ProtheusDenuo = denuo;
            _WEB = wEB;
        }

        public IActionResult OnGet(int id)
        {
            Agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == id);
            var data = Convert.ToInt32(DateTime.Now.ToString("yyyy/MM/dd").Replace("/", ""));

            var codigos = SGID.ProdutosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id)
                    .Select(x => new
                    {
                        produto = x.CodigoProduto,
                        unidade = x.Quantidade,
                        valorunitario = x.ValorUnitario,
                        valor = x.ValorTotal,
                        Tabela = x.CodigoTabela
                    }).ToList();

            Produtos = new List<Produto>();

            if (Agendamento.Empresa == "01")
            {
                //Intermedic

                codigos.ForEach(x =>
                {
                    var produto = (from SB10 in ProtheusInter.Sb1010s
                                   where SB10.B1Cod == x.produto && SB10.B1Msblql != "1"
                                   && SB10.DELET != "*"
                                   select new
                                   {
                                       SB10.B1Cod,
                                       SB10.B1Msblql,
                                       SB10.B1Solicit,
                                       SB10.B1Desc,
                                       SB10.B1Fabric,
                                       SB10.B1Tipo,
                                       SB10.B1Lotesbp,
                                       SB10.B1Um,
                                       SB10.B1Reganvi,
                                       SB10.B1Xtuss
                                   }).FirstOrDefault();

                    var ViewProduto = new Produto
                    {
                        Item = produto.B1Cod,
                        Licit = produto.B1Solicit,
                        Produtos = produto.B1Desc,
                        Tuss = produto.B1Xtuss,
                        Anvisa = produto.B1Reganvi,
                        Marca = produto.B1Fabric,
                        //Validade = produto.B8Dtvalid,
                        Und = x.unidade,
                        PrcUnid = x.valorunitario,
                        SegUnd = produto.B1Um,
                        //Descon = produto.Da1Vlrdes,
                        VlrTotal = x.valor,
                        TipoOp = produto.B1Tipo,
                    };

                    UnidadesTotal += x.unidade;
                    ValorTotal += x.valor;

                    Produtos.Add(ViewProduto);
                });
            }
            else
            {
                //Denuo
                codigos.ForEach(x =>
                {
                    var produto = (from SB10 in ProtheusDenuo.Sb1010s
                                   where SB10.B1Cod == x.produto && SB10.B1Msblql != "1"
                                   && SB10.DELET != "*"
                                   select new
                                   {
                                       SB10.B1Cod,
                                       SB10.B1Msblql,
                                       SB10.B1Solicit,
                                       SB10.B1Desc,
                                       SB10.B1Fabric,
                                       SB10.B1Tipo,
                                       SB10.B1Lotesbp,
                                       SB10.B1Um,
                                       SB10.B1Reganvi,
                                       SB10.B1Xtuss
                                   }).FirstOrDefault();

                    var ViewProduto = new Produto
                    {
                        Item = produto.B1Cod,
                        Licit = produto.B1Solicit,
                        Produtos = produto.B1Desc,
                        Tuss = produto.B1Xtuss,
                        Anvisa = produto.B1Reganvi,
                        //Validade = produto.B8Dtvalid,
                        Marca = produto.B1Fabric,
                        Und = x.unidade,
                        PrcUnid = x.valorunitario,
                        SegUnd = produto.B1Um,
                        //Descon = produto.Da1Vlrdes,
                        VlrTotal = x.valor,
                        TipoOp = produto.B1Tipo,
                    };

                    UnidadesTotal += x.unidade;
                    ValorTotal += x.valor;

                    Produtos.Add(ViewProduto);
                });
            }

            string Pasta = $"{_WEB.WebRootPath}/Rotativa/Rotativa";

            if (Agendamento.StatusPedido != 3 && Agendamento.StatusPedido != 4) 
            {
                Agendamento.StatusPedido = 6;
                SGID.Agendamentos.Update(Agendamento);
                SGID.SaveChanges();
            }
            
            return Page();
        }
    }
}
