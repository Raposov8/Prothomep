using DocumentFormat.OpenXml.Bibliography;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Denuo;
using SGID.Models.DTO;
using SGID.Models.Estoque;

namespace SGID.Pages.Estoque
{
    [Authorize]
	public class FormularioAvulsoModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private TOTVSDENUOContext ProtheusDenuo { get; set; }

        public string Empresa { get; set; }

        public List<string> Codigos { get; set; }

        public FormularioAvulsoModel(ApplicationDbContext sgid,TOTVSINTERContext inter,TOTVSDENUOContext denuo)
        {
            SGID = sgid;
            ProtheusInter = inter;
            ProtheusDenuo = denuo;
        }
        public void OnGet(string id)
        {
            Empresa = id;

            if(Empresa == "01")
            {
                Codigos = ProtheusInter.Sb1010s.Where(x => x.B1Msblql != "1" && x.DELET != "*")
                    .Select(x => x.B1Cod).ToList();
            }
            else
            {
                Codigos = ProtheusDenuo.Sb1010s.Where(x => x.B1Msblql != "1" && x.DELET != "*")
                    .Select(x => x.B1Cod ).ToList();
            }
        }

        public IActionResult OnPost(string Cliente,DateTime DataCirurgia,string Paciente,string Medico,
            string Representante, string NumAgendamento,string Convenio,List<Produto> Produtos,string Empresa)
        {
            var Formulario = new FormularioAvulso
            {
                DataCriacao = DateTime.Now,

                Cliente = Cliente,
                DataCirurgia = DataCirurgia,
                Paciente = Paciente,
                Cirurgiao = Medico,
                Convenio = Convenio,
                Representante = Representante,
                NumAgendamento = NumAgendamento,
                Usuario = User.Identity.Name.Split("@")[0],
                Empresa = Empresa
            };

            SGID.FormularioAvulsos.Add(Formulario);
            SGID.SaveChanges();

            Produtos.ForEach(produto =>
            {
                var ProdutoForm = new FormularioAvulsoXProdutos
                {
                    FormularioId = Formulario.Id,
                    Produto = produto.Item,
                    Descricao = produto.Produtos,
                    Quantidade = produto.Und,
                    Lote = produto.Lote
                };

                SGID.FormularioAvulsoXProdutos.Add(ProdutoForm);
                SGID.SaveChanges();
            });



            return LocalRedirect($"/estoque/formularioavulsopdf/{Formulario.Id}");
        }

        public JsonResult OnPostAdicionarProd(string Codigo, string Empresa)
        {
            try
            {
                if (Empresa == "01")
                {

                    //Intermedic
                    var produto = (from SB10 in ProtheusInter.Sb1010s
                                   where SB10.B1Cod == Codigo.ToUpper() && SB10.B1Msblql != "1"
                                   && SB10.DELET != "*"
                                   select new
                                   {
                                       SB10.B1Cod,
                                       SB10.B1Msblql,
                                       SB10.B1Solicit,
                                       SB10.B1Desc,
                                   }).FirstOrDefault();

                    if (produto != null)
                    {

                            var ViewProduto = new Produto
                            {
                                Item = produto.B1Cod,
                                Produtos = produto.B1Desc,
                                Und = 0,
                            };

                            return new JsonResult(ViewProduto);
                        
                    }
                }
                else
                {
                    //Denuo

                    var produto = (from SB10 in ProtheusDenuo.Sb1010s
                                   where SB10.B1Cod == Codigo.ToUpper() && SB10.B1Msblql != "1"
                                   && SB10.DELET != "*"
                                   select new
                                   {
                                       SB10.B1Cod,
                                       SB10.B1Msblql,
                                       SB10.B1Solicit,
                                       SB10.B1Desc,
                                   }).FirstOrDefault();

                    if (produto != null)
                    {

                        var ViewProduto = new Produto
                        {
                            Item = produto.B1Cod,
                            Produtos = produto.B1Desc,
                            Und = 0,
                        };

                        return new JsonResult(ViewProduto);
                    }
                }
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "FormularioAvulso Adicionar", user);
            }

            return new JsonResult("");
        }
    }
}
