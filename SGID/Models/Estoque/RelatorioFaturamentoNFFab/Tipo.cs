namespace SGID.Models.Estoque.RelatorioFaturamentoNFFab
{
    public class Tipo
    {
        public string Nome { get; set; }
        public string NomeFabricante { get; set; }
        public double Janeiro { get; set; }
        public double Fevereiro { get; set; }
        public double Marco { get; set; }
        public double Abril { get; set; }
        public double Maio { get; set; }
        public double Junho { get; set; }
        public double Julho { get; set; }
        public double Agosto { get; set; }
        public double Setembro { get; set; }
        public double Outubro { get; set; }
        public double Novembro { get; set; }
        public double Dezembro { get; set; }
        public double Total { get; set; }
        public List<Cliente> Clientes { get; set; } = new List<Cliente>();
    }
}
