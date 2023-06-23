using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class FormularioAvulso
    {
        [Key]
        public int Id { get; set; }
        public DateTime DataCriacao { get; set; }

        public string Cliente { get; set; }
        public DateTime DataCirurgia { get; set; }
        public string Paciente { get; set; }
        public string Cirurgiao { get; set; }
        public string Convenio { get; set; }
        public string Representante { get; set; }
        public string NumAgendamento { get; set; }
        public string Usuario { get; set; }
        public string Empresa { get; set; }
    }
}
