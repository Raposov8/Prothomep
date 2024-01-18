using System.IO.Packaging;

namespace SGID.Models.AdmVendas
{
    public class RelatorioCirurgiaAFaturar
    {
        public string Vendedor { get; set; }
        public string Cliente { get; set; }
        public string GrupoCliente { get; set; }
        public string ClienteEntrega { get; set; }
        public string Medico { get; set; }
        public string Convenio { get; set; }
        public string Cirurgia { get; set; }
        public string Hoje { get; set; }
        public int Dias { get; set; }
        public string Anging { get; set; }
        public string Paciente { get; set; }
        public string Matric { get; set; }
        public string INPART { get; set; }
        public double Valor { get; set; }
        public string PVFaturamento { get; set; }
        public string Status { get; set; }
        public string DataEnvRA { get; set; }
        public string DataRecRA { get; set; }
        public string DataValorizacao { get; set; }
        public string DataEmissao { get; set; }
    }
}
