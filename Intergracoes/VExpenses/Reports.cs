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
    public class Reports
    {
        private string IdApi { get; set; } = "74pvl26wkUaIiLSl2eV6X4GDeLQCcayjtzJRvlhUPTvLT0xpswrzHlgRgn0N";

        private string Url { get; set; } = "https://api.vexpenses.com/v2";

        public async Task<List<Reportar>> GetReportes()
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.GetAsync($"{Url}/reports");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<List<Reportar>>(texto);
        }

        public async Task<Reportar> GetReporteById(int Id)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.GetAsync($"{Url}/reports/{Id}");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Reportar>(texto);
        }

        public async Task<Reportar> PostReportar(Reportar projeto)
        {
            var client = new HttpClient();
            var content = new StringContent(JsonConvert.SerializeObject(projeto), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.PostAsync($"{Url}/reports/", content);
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Reportar>(texto);
        }

        public async Task<Reportar> PutReportar(Reportar projeto)
        {
            var client = new HttpClient();
            var content = new StringContent(JsonConvert.SerializeObject(projeto), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.PutAsync($"{Url}/reports/{projeto.id}", content);
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Reportar>(texto);
        }

        /*public async Task<Projeto> PostApprove(int Id)
        {
            var client = new HttpClient();
            var content = new StringContent(JsonConvert.SerializeObject(projeto), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.PostAsync($"{Url}/projects/{Id}");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Custo>(texto);
        }*/
    }
}
