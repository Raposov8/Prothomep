namespace SGID.Models.SubDistribuidor
{
    public class RelatorioSubDistribuidorPorFabricante
    {
        public string Codigo { get; set; }
        public string Nome { get; set; }
        public string NF { get; set; }
        public string Emissao { get; set; }
        public string Fabricante { get; set; }
        public string CodProduto { get; set; }
        public string Descricao { get; set; }
        public double Quant { get; set; }
        public double PrcVen { get; set; }
        public double Total { get; set; }
        public string NomeVendedor { get; set; }
    }
}
