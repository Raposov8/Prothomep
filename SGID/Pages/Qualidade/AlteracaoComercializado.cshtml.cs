using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;
using SGID.Models.Inter;

namespace SGID.Pages.Qualidade
{
    public class AlteracaoComercializadoModel : PageModel
    {
        private TOTVSDENUOContext ProtheusDenuo { get; set; }
        private TOTVSINTERContext ProtheusInter { get; set; }
        private ApplicationDbContext SGID { get; set; }

        private readonly IWebHostEnvironment _WEB;

        public string TextoResposta { get; set; } = "";

        public AlteracaoComercializadoModel(TOTVSDENUOContext denuo, TOTVSINTERContext inter, ApplicationDbContext sgid, IWebHostEnvironment web)
        {
            ProtheusDenuo = denuo;
            ProtheusInter = inter;
            SGID = sgid;
            _WEB = web;
        }

        public void OnGet()
        {

        }



        public IActionResult OnPost(IFormCollection Anexos,string Empresa)
        {
            var produto = "";
            var campo = "";
            try
            {
                string Pasta = $"{_WEB.WebRootPath}/Temp";

                var Data = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss").Replace("/", "").Replace(" ", "").Replace(":", "");

                if (!Directory.Exists(Pasta))
                {
                    Directory.CreateDirectory(Pasta);
                }

                if(Empresa == "DENUO")
                {
                    foreach (var anexo in Anexos.Files)
                    {
                        string Caminho = $"{Pasta}/AlteracaoProduto{Data}.csv";

                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            anexo.CopyTo(fileStream);
                        }
                    }

                    using (var reader = new StreamReader($"{Pasta}/AlteracaoProduto{Data}.csv"))
                    {
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();

                            if (line.Contains(';'))
                            {
                                var values = line.Split(';');

                                produto = values[0];

                                var DBProduto = ProtheusDenuo.Sb1010s.FirstOrDefault(x => x.B1Cod == produto);

                                DBProduto.B1Comerci = values[1];

                                ProtheusDenuo.Update(DBProduto);
                                ProtheusDenuo.SaveChanges();

                            }
                        }
                    }
                }
                else
                {
                    foreach (var anexo in Anexos.Files)
                    {
                        string Caminho = $"{Pasta}/AlteracaoProdutoInter{Data}.csv";

                        using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                        {
                            anexo.CopyTo(fileStream);
                        }
                    }

                    using (var reader = new StreamReader($"{Pasta}/AlteracaoProdutoInter{Data}.csv"))
                    {
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();

                            if (line.Contains(';'))
                            {
                                var values = line.Split(';');

                                produto = values[0];

                                
                                var DBProduto = ProtheusInter.Sb1010s.FirstOrDefault(x => x.B1Cod == produto);

                                DBProduto.B1Comerci = values[1];

                                ProtheusInter.Update(DBProduto);
                                ProtheusInter.SaveChanges();
                                
                            }
                        }
                    }
                }

                

                TextoResposta = $"Informações Mudadas com Sucesso";
            }
            catch (FormatException e)
            {

                TextoResposta = $"Error: Checar a formatação do Texto Produto:{produto}";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar Produto Denuo", user);

                return Page();
            }
            catch (ArgumentNullException e)
            {
                TextoResposta = $"Error: Não encotrado o produto:{produto}";
               
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar Produto Denuo", user);

                return Page();
            }
            catch (DbUpdateException e)
            {

                if (e.InnerException != null)
                {
                    TextoResposta = $"Produto: {produto} Error: {e.InnerException.Message}";
                }
                else
                {
                    TextoResposta = $"Produto: {produto} Error: {e.Message}";
                }

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Comercializado Produto Denuo", user);

                return Page();
            }
            catch (Exception e)
            {
               

                TextoResposta = "Error: Contate o TI";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar Produto Denuo", user);

                return Page();
            }

            return Page();
        }
    }
}
