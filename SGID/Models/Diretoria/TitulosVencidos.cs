namespace SGID.Models.Diretoria
{
    public class TitulosVencidos
    {
        public string CodCliente { get; set; }
        public string NF { get; set; }
        public string Cliente { get; set; }
        public double Valor { get; set; }
        public double ValorSaldo { get; set; }

        public string DataEmissao { get; set; }
        public string DataVencimento { get; set; }
    }
}
