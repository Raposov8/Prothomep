namespace SGID.Models.Gerencial.RelatorioCirurgiaNFaturada
{
    public class Operacao
    {
        public string Emissao { get; set; }
        public string Numero { get; set; }
        public string Nome { get; set; }
        public string Paciente { get; set; }
        public string Medico { get; set; }
        public string TipoOper { get; set; }
        public string TipoCli { get; set; }
        public string UPatrim { get; set; }
        public string DataCirurgia { get; set; }
        public double Dias { get; set; }
        public string Periodo { get; set; }
        public string Order { get;set; }
        public double Valor { get; set; }
    }
}
