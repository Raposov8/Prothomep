namespace SGID.Models.Financeiro
{
    public class RelatorioAPagarMes
    {
        public string Empresa { get; set; }
        public string NumDia { get; set; }
        public string DiaSem { get; set; }
        public string Semana { get; set; }
        public string AnoMes { get; set; }
        public string E2Naturez { get; set; }
        public string EdDescric { get; set; }
        public string EdXgrdesp { get; set; }
        public string GRPDESC { get; set; }
        public string E2Fornece { get; set; }
        public string E2Loja { get; set; }
        public string E2Nomfor { get; set; }
        public string E2Vencrea { get; set; }
        public string E2Hist { get; set; }
        public double VLPagar { get; set; }
        public double VLPrPagar { get; set; }
        public double VLOrig { get; set; }
        public string Tipo { get; set; }
    }
}
