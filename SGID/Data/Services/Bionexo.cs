using Newtonsoft.Json;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using OPMEnexo;

namespace SGID.Data.Services
{
    public class Bionexo
    {
        //public string Url { get; set; } = "https://sandbox.opmenexo.com.br/ws/WSPedido";

        public string Email { get; set; } = "ALINE01";

        public string Password { get; set; } = "Davi*2808";


        public WSPedidoClient Web = new WSPedidoClient();

        public async Task<object> IncluirPedido(pedido pedido)
        {
            var usuario = new Usuario { login = Email, senha = Password };
            var resp = await Web.incluirPedidoAsync(usuario,pedido);
            //var resp = await Web.

            return usuario;
        }
    }
}
