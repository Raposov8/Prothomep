namespace SGID.Models.Financeiro
{
    public class ReceitaBruta
    {
        public string Produto { get; set; }
        public string Descricao { get; set; }
        public string NFSaida { get; set; }
        public string CFOP { get; set; }
        public string Cliente { get; set; }
        public string Tipo { get; set; }
        public double Bruta { get; set; }
        public double Imposto { get; set; }
        public double Qtde { get; set; }
        public string Parcelas { get; set; }
        public string CondPag { get; set; }
        public double Custo { get; set; }
        public string Data { get; set; }
        public string Pedido { get; set; }
        public string Vendedor { get; set; }
        public string Fabricante { get; set; }
    }
}
