using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Data.ViewModel;
using SGID.Models.Cirurgias;
using SGID.Models.Denuo;
using SGID.Models.DTO;
using SGID.Models.Inter;

namespace SGID.Pages.Logistica
{
    [Authorize]
    public class PickListLogisticaModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private readonly IWebHostEnvironment _WEB;
        public Agendamentos Agendamento { get; set; }

        public List<Produto> Produtos { get; set; } = new List<Produto>();
        public List<Patrimonio> Patrimonios { get; set; } = new List<Patrimonio>();
        public List<Produto> Avulsos { get; set; } = new List<Produto>();
        public double UnidadesTotal { get; set; }
        public double ValorTotal { get; set; }

        public PickListLogisticaModel(ApplicationDbContext sGID, TOTVSDENUOContext denuo, TOTVSINTERContext inter, IWebHostEnvironment wEB)
        {
            SGID = sGID;
            ProtheusInter = inter;
            ProtheusDenuo = denuo;
            _WEB = wEB;
        }

        public IActionResult OnGet(int id)
        {
            Agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == id);

            var codigos = SGID.ProdutosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id)
                .Select(x => new
                {
                    produto = x.CodigoProduto,
                    unidade = x.Quantidade,
                    valor = x.ValorTotal
                }).ToList();

            var patrimonios = SGID.PatrimoniosAgendamentos.Where(x => x.AgendamentoId == Agendamento.Id).ToList();

            var avulsos = SGID.AvulsosAgendamento.Where(x => x.AgendamentoId == Agendamento.Id).ToList();

            if (Agendamento.Empresa == "01")
            {
                codigos.ForEach(x =>
                {
                    var produto = ProtheusInter.Sb1010s.Where(c => c.B1Cod == x.produto).Select(c => new Produto
                    {
                        Item = c.B1Cod,
                        Licit = c.B1Solicit,
                        Produtos = c.B1Desc,
                        Tuss = c.B1Xtuss,
                        Anvisa = c.B1Reganvi,
                        Marca = c.B1Fabric,
                        Und = x.unidade,
                        PrcUnid = c.B1Uprc,
                        SegUnd = c.B1Um,
                        VlrTotal = x.valor,
                        TipoOp = c.B1Tipo,
                    }).First();

                    Produtos.Add(produto);
                });

                patrimonios.ForEach(x =>
                {
                    var view = (from PA10 in ProtheusInter.Pa1010s
                                join PAC in ProtheusInter.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                from c in sr.DefaultIfEmpty()
                                join SA10 in ProtheusInter.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                                from a in st.DefaultIfEmpty()
                                where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                                && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == x.Patrimonio
                                select new Patrimonio
                                {
                                    Descri = PA10.Pa1Despat,
                                    KitBas = PA10.Pa1Kitbas,
                                    Quantidade = 1
                                }).First();


                    Patrimonios.Add(view);
                });

                avulsos.ForEach(x =>
                {
                    var avulso = ProtheusInter.Sb1010s.Where(c => c.B1Cod == x.Produto).Select(c => new Produto
                    {
                        Item = c.B1Cod,
                        Licit = c.B1Solicit,
                        Produtos = c.B1Desc,
                        Tuss = c.B1Xtuss,
                        Anvisa = c.B1Reganvi,
                        Marca = c.B1Fabric,
                        Und = x.Quantidade,
                        TipoOp = c.B1Tipo,
                    }).First();

                    Avulsos.Add(avulso);

                });
            }
            else
            {
                codigos.ForEach(x =>
                {
                    var produto = ProtheusDenuo.Sb1010s.Where(c => c.B1Cod == x.produto).Select(c => new Produto
                    {
                        Item = c.B1Cod,
                        Licit = c.B1Solicit,
                        Produtos = c.B1Desc,
                        Tuss = c.B1Xtuss,
                        Anvisa = c.B1Reganvi,
                        Marca = c.B1Fabric,
                        Und = x.unidade,
                        PrcUnid = c.B1Uprc,
                        SegUnd = c.B1Um,
                        VlrTotal = x.valor,
                    }).First();

                    Produtos.Add(produto);
                });

                patrimonios.ForEach(x =>
                {
                    var view = (from PA10 in ProtheusDenuo.Pa1010s
                                join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                from c in sr.DefaultIfEmpty()
                                join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                                from a in st.DefaultIfEmpty()
                                where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                                && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == x.Patrimonio
                                select new Patrimonio
                                {
                                    Descri = PA10.Pa1Despat,
                                    KitBas = PA10.Pa1Kitbas,
                                    Quantidade = 1
                                }).First();


                    Patrimonios.Add(view);
                });

                avulsos.ForEach(x =>
                {
                    var avulso = ProtheusDenuo.Sb1010s.Where(c => c.B1Cod == x.Produto).Select(c => new Produto
                    {
                        Item = c.B1Cod,
                        Licit = c.B1Solicit,
                        Produtos = c.B1Desc,
                        Tuss = c.B1Xtuss,
                        Anvisa = c.B1Reganvi,
                        Marca = c.B1Fabric,
                        Und = x.Quantidade,
                        TipoOp = c.B1Tipo,
                    }).First();

                    Avulsos.Add(avulso);

                });
            }

            return Page();
        }
    }
}
