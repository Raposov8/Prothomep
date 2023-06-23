using Intergracoes.VExpenses.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses
{
    public class CostsCenters
    {
        private string IdApi { get; set; } = "74pvl26wkUaIiLSl2eV6X4GDeLQCcayjtzJRvlhUPTvLT0xpswrzHlgRgn0N";

        private string Url { get; set; } = "https://api.vexpenses.com/v2";

        public async Task<List<Custo>> GetCustos()
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization",$"{IdApi}"); //Chave
            var resultPost = await client.GetAsync($"{Url}/costs-centers");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<List<Custo>>(texto);
        }

        public async Task<Custo> GetCustoById(int Id)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.GetAsync($"{Url}/costs-centers/{Id}");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Custo>(texto);
        }

        public async Task<Custo> PostCusto(Custo custo)
        {
            var client = new HttpClient();
            var content = new StringContent(JsonConvert.SerializeObject(custo), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.PostAsync($"{Url}/costs-centers/",content);
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Custo>(texto);
        }

        public async Task<Custo> PutCusto(Custo custo)
        {
            var client = new HttpClient();
            var content = new StringContent(JsonConvert.SerializeObject(custo), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.PutAsync($"{Url}/costs-centers/{custo.id}", content);
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Custo>(texto);
        }

        public async Task<Custo> DeleteCusto(int Id)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.DeleteAsync($"{Url}/costs-centers/{Id}");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Custo>(texto);
        }
    }
}
