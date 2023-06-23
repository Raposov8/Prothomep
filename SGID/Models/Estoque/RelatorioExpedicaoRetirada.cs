namespace SGID.Models.Estoque
{
    public class RelatorioExpedicaoRetirada
    {
        public string Agendamento { get; set; }
        public string AgPrincipal { get; set; }
        public string CodCliente { get; set; }
        public string LojaCliente { get; set; }
        public string Nome { get; set; }
        public string TipoOper { get; set; }
        public string DTCirurgia { get; set; }
        public string HoraCirurgia { get; set; }
        public string Status { get; set; }
        public string CodCanc { get; set; }
        public string MotivoCanc { get; set; }
        public string DTExpedicao { get; set; }
        public string HRExpedicao { get; set; }
        public string MotEntrega { get; set; }
        public string DTRetirada { get; set; }
        public string HRRetirada { get; set; }
        public string MotRetirada { get; set; }
    }
}
