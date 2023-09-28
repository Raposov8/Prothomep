namespace SGID.Models.Qualidade
{
    public class RelatorioValidadePermanente
    {
        public string Codigo { get; set; }
        public string Descricao { get; set; }
        public string Filial { get; set; }
        public string Local { get; set; }
        public string Lote { get; set; }
        public string LoteFor { get; set; }
        
        public string ValidLote { get; set; }
        public double Saldo { get; set; }
    }
}
