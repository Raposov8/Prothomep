namespace SGID.Models.Comercial
{
    public class GestorComercialDash
    {
        public string Nome { get; set; }
        public string User { get; set; }
        public int CirurgiasValorizadas { get; set; }
        public double CirurgiasValorizadasValor { get; set; }
        public int FaturadoMes { get; set; }
        public double FaturadoMesValor { get; set; }
        public int CirurgiasEmAberto { get; set; }
        public double CirurgiasEmAbertoValor { get; set; }
        public int CirurgiasLicitacoes { get; set; }
        public double CirurgiasLicitacoesValor { get; set; }
        public double Meta { get; set; }
    }
}
