using Intergracoes.Inpart.Models;
using Intergracoes.VExpenses;
using Intergracoes.VExpenses.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.Inpart
{
    public class IntegracaoInPart
    {
        private string BaseUrl { get; set; } = "https://integracao-api.inpartsaudeteste.com.br/api";
        private string Auth { get; set; } = "";

        public async Task<string> Authenticacao()
        {

            var credenciais = new 
            { 
                usuario = "intermedicsp.integracao",
                senha = "intermedicsp@28319"
            };


            var client = new HttpClient();

            var content = new StringContent(JsonConvert.SerializeObject(credenciais), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

            var resultPost = await client.PostAsync($"{BaseUrl}/",content);
            var data = JsonConvert.DeserializeObject<Token>(resultPost.Content.ReadAsStringAsync().Result);
           
            return data.AccessToken;

        }


        public async Task<List<Cotacao>> ListarCotacoes()
        {

            var client = new HttpClient();

            var credenciais = Authenticacao();

            client.DefaultRequestHeaders.Add("Authorization", $"Bearer {credenciais}");

            var resultPost = await client.GetAsync($"{BaseUrl}/cotacao");
            var data = JsonConvert.DeserializeObject<List<Cotacao>>(resultPost.Content.ReadAsStringAsync().Result);

            return data;
        }
    }
}
