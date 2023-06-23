using Intergracoes.VExpenses.Despecas;
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
    public class Expenses
    {
        private string IdApi { get; set; } = "74pvl26wkUaIiLSl2eV6X4GDeLQCcayjtzJRvlhUPTvLT0xpswrzHlgRgn0N";

        private string Url { get; set; } = "https://api.vexpenses.com/v2";

        public async Task<List<Despecas.Despecas>> GetDespecas(DateTime Inicio, DateTime Fim)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.GetAsync($"{Url}/expenses?include=user,expense_type,costs_center,report&search=date:{Inicio:yyyy-MM-dd},{Fim:yyyy-MM-dd}&searchFields=date:between");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<List<Despecas.Despecas>>(texto);
        }

        public async Task<Despecas.Despecas> GetDespecasById(int Id)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.GetAsync($"{Url}/expenses/{Id}");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Despecas.Despecas>(texto);
        }

        public async Task<Despecas.Despecas> PostDespecas(Despecas.Despecas projeto)
        {
            var client = new HttpClient();
            var content = new StringContent(JsonConvert.SerializeObject(projeto), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.PostAsync($"{Url}/expenses/", content);
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Despecas.Despecas>(texto);
        }

        public async Task<Despecas.Despecas> PutDespecas(Despecas.Despecas projeto)
        {
            var client = new HttpClient();
            var content = new StringContent(JsonConvert.SerializeObject(projeto), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.PutAsync($"{Url}/expenses/{projeto.id}", content);
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Despecas.Despecas>(texto);
        }

        public async Task<Despecas.Despecas> DeleteDespecas(int Id)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.DeleteAsync($"{Url}/projects/{Id}");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Despecas.Despecas>(texto);
        }
    }
}
