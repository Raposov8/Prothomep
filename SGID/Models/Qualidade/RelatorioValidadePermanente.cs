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
        public string Processo { get; set; }
        public string Agendamento { get; set; }
        public string Patrimonio { get; set; }
        public string Operacao { get; set; }
        public string NFSaida { get; set; }
        public string SerieSaida { get; set; }
        public string EmissaoNf { get; set; }
        public string ItSaida { get; set; }
        public double QtdSaida { get; set; }
        public double Saldo { get; set; }
    }
}
