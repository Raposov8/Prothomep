using SGID.Data.ViewModel;

namespace SGID.Models.Account.RH
{
    public class TimeDentalRH
    {
        public TimeDental User { get; set; }
        public string Nome { get; set; }
        public double Faturado { get; set; }
        public double Comissao { get; set; }
        public double FaturadoEquipe { get; set; }
        public double ComissaoEquipe { get; set; }

        public double Meta { get; set; }
        public double MetaAtingimento { get; set; }
        public double Salario { get; set; }
        public double Garantia { get; set; }

        public double Total { get; set; }
        public double Teto { get; set; }
        public double Paga { get; set; }
        public string Linha { get; set; }
    }
}
