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
    public class Projects
    {
        private string IdApi { get; set; } = "74pvl26wkUaIiLSl2eV6X4GDeLQCcayjtzJRvlhUPTvLT0xpswrzHlgRgn0N";

        private string Url { get; set; } = "https://api.vexpenses.com/v2";

        public async Task<List<Projeto>> GetProjetos()
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.GetAsync($"{Url}/projects");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<List<Projeto>>(texto);
        }

        public async Task<Projeto> GetProjetoById(int Id)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.GetAsync($"{Url}/projects/{Id}");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Projeto>(texto);
        }

        public async Task<Projeto> PostProjeto(Projeto projeto)
        {
            var client = new HttpClient();
            var content = new StringContent(JsonConvert.SerializeObject(projeto), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.PostAsync($"{Url}/projects/", content);
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Projeto>(texto);
        }

        public async Task<Projeto> PutProjeto(Projeto projeto)
        {
            var client = new HttpClient();
            var content = new StringContent(JsonConvert.SerializeObject(projeto), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.PutAsync($"{Url}/projects/{projeto.id}", content);
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Projeto>(texto);
        }

        public async Task<Projeto> DeleteProjeto(int Id)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.DeleteAsync($"{Url}/projects/{Id}");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<Projeto>(texto);
        }
    }
}
