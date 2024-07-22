using DocumentFormat.OpenXml.Office.CoverPageProps;

namespace SGID.Models.Patrimonio
{
    public class Patrimonio
    {
        public string CodProd { get; set; }
        public string DescProd { get; set; }
        public double QtdKit { get; set; }
        public double QtdPat { get; set; }

        //Itens Faltantes
        public string CodPatri { get; set; }
        public double QuantFalt { get; set; }

        public string DescPatri { get; set; }
        public string Bloqueio { get; set; }
    }
}
