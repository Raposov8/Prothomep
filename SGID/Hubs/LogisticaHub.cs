using Microsoft.AspNetCore.SignalR;

namespace SGID.Hubs
{
    public class LogisticaHub : Hub
    {
        public async Task SendMessage(string User,string Mensagem)
        {
            await Clients.All.SendAsync("ReceiveMessage", User, Mensagem);
        }
    }
}
