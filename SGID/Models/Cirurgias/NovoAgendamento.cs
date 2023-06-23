

using SGID.Data.ViewModel;

namespace SGID.Models.Cirurgias
{
    public class NovoAgendamento
    {
        public List<string> Clientes { get; set; } = new List<string>();
        public List<string> Convenio { get; set; } = new List<string>();
        public List<string> Medico { get; set; } = new List<string>();
        public List<string> Intrumentador { get; set; } = new List<string>();
        public List<string> Hospital { get; set; } = new List<string>();
        public List<Procedimento> Procedimentos { get; set; } = new List<Procedimento>();
        public List<string> Patrimonios { get; set; } = new List<string>();
    }
}
