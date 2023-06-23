using SGID.Data.ViewModel;

namespace SGID.Models.Account.RH
{
    public class TimeRH
    {

        public Time User { get; set; }
        public string Nome { get; set; }
        public double Faturado { get; set; }
        public double Comissao { get; set; }
        public double FaturadoEquipe { get; set; }
        public double ComissaoEquipe { get; set; }
        public double FaturadoProduto { get; set; }
        public double ComissaoProduto { get; set; }
        public string Linha { get; set; }
    }
}
