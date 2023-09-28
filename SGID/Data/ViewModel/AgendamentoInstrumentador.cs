using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class AgendamentoInstrumentador
    {
        [Key]
        public int Id { get; set; }
        public string Empresa { get; set; }
        public string Data { get; set; }
        public string Instrumentador { get; set; }
        public string Agendamento { get; set; }
        public string Procedimento { get; set; }
        public string Hospital { get; set; }
        public string Medico { get; set; }
        public string Tipo { get; set; }
        public string Status { get; set; }
    }
}
