namespace SGID.Models.Estoque
{
    public class RelatorioGiroEstoque
    {
        public string Codigo { get; set; }
        public string Fabricante { get; set; }
        public string Unidade { get; set; }
        public string Desc { get; set; }
        public double Saldo { get; set; }
        public double Saida { get; set; }
        public double Giro { get; set; }
        public double EstoqueMinimo { get; set; }
        public double PontoPedido { get; set; }
    }
}
