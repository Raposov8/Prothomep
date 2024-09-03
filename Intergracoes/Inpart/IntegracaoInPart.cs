using Intergracoes.Inpart.Models;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Text;

namespace Intergracoes.Inpart
{
    public class IntegracaoInPart
    {
        public string BaseUrl { get; set; } = @"https://integracao-api.inpartsaude.com.br/api";
        public string Auth { get; set; } = "";

        public async Task<string> Authenticacao(string Empresa)
        {
            try
            {
                if(Empresa == "01")
                {
                    var credenciais = new
                    {
                        usuario = "Rogeriomartines.i",
                        senha = "Roger@2024"
                    };

                    var client = new HttpClient();

                    var content = new StringContent(JsonConvert.SerializeObject(credenciais), Encoding.UTF8, "application/json");
                    content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

                    var resultPost = await client.PostAsync($"{BaseUrl}/autenticacao", content);
                    var data = JsonConvert.DeserializeObject<Token>(resultPost.Content.ReadAsStringAsync().Result);

                    return data.AccessToken;
                }
                else
                {
                    var credenciais = new
                    {
                        usuario = "integracaodenuo.d",
                        senha = "Roger@2024"
                    };


                    var client = new HttpClient();

                    var content = new StringContent(JsonConvert.SerializeObject(credenciais), Encoding.UTF8, "application/json");
                    content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

                    var resultPost = await client.PostAsync($"{BaseUrl}/autenticacao", content);
                    var data = JsonConvert.DeserializeObject<Token>(resultPost.Content.ReadAsStringAsync().Result);

                    return data.AccessToken;
                }

            }
            catch (Exception e)
            {
                var email = e.Message;
                return "";
            }

        }


        public async Task<List<Cotacao>> ListarCotacoes(string Empresa,DateTime DtCotacaoInicio,DateTime DtCotacaoFim)
        {

            var client = new HttpClient();

            var credenciais = await Authenticacao(Empresa);

            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {credenciais}");

            var query = "";

            if(DtCotacaoInicio != new DateTime(1, 1, 1) && DtCotacaoFim!= new DateTime(1, 1, 1))
            {
                query = $"DtCotacaoInicial={DtCotacaoInicio}&DtCotacaoFinal={DtCotacaoFim}";
            }


            var resultPost = await client.GetAsync($"{BaseUrl}/cotacao?{query}");
            var data = JsonConvert.DeserializeObject<List<Cotacao>>(resultPost.Content.ReadAsStringAsync().Result);

            return data;
        }
    }
}
