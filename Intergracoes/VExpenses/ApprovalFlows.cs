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
    public class ApprovalFlows
    {
        private string IdApi { get; set; } = "74pvl26wkUaIiLSl2eV6X4GDeLQCcayjtzJRvlhUPTvLT0xpswrzHlgRgn0N";

        private string Url { get; set; } = "https://api.vexpenses.com/v2";

        public async Task<List<FluxoAprovacao>> GetFluxos()
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.GetAsync($"{Url}/approval-flows");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<List<FluxoAprovacao>>(texto);
        }

        public async Task<FluxoAprovacao> GetFluxoById(int Id)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.GetAsync($"{Url}/approval-flows/{Id}");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<FluxoAprovacao>(texto);
        }

        public async Task<FluxoAprovacao> PostFluxo(FluxoAprovacao projeto)
        {
            var client = new HttpClient();
            var content = new StringContent(JsonConvert.SerializeObject(projeto), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.PostAsync($"{Url}/approval-flows/", content);
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<FluxoAprovacao>(texto);
        }

        public async Task<FluxoAprovacao> PutFluxo(FluxoAprovacao projeto)
        {
            var client = new HttpClient();
            var content = new StringContent(JsonConvert.SerializeObject(projeto), Encoding.UTF8, "application/json");
            content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.PutAsync($"{Url}/approval-flows/{projeto.id}", content);
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<FluxoAprovacao>(texto);
        }

        public async Task<FluxoAprovacao> DeleteFluxo(int Id)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", $"{IdApi}"); //Chave
            var resultPost = await client.DeleteAsync($"{Url}/approval-flows/{Id}");
            var data = JsonConvert.DeserializeObject<dynamic>(resultPost.Content.ReadAsStringAsync().Result);
            var texto = JsonConvert.SerializeObject(data.data);
            return JsonConvert.DeserializeObject<FluxoAprovacao>(texto);
        }
    }
}
