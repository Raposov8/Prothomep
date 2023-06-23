namespace SGID.Models.Estoque
{
    public class RelatorioSaldoEstoque
    {
        public string Filial { get; set; }
        public string Produto { get; set; }
        public string DescProd { get; set; }
        public string LoteCtl { get; set; }
        public string Local { get; set; }
        public string Endereco { get; set; }
        public double Saldo { get; set; }
        public string UM { get; set; }
        public int Empenho { get;set; }
        public string Data { get; set; }
        public string DtValid { get; set; }
        public string Descendereco { get; set; }
    }
}
