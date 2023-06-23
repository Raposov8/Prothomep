using System.ComponentModel.DataAnnotations;

namespace SGID.Models.DTO
{
    public class Cirurgia
    {
        public int Id { get; set; }
        public DateTime DataCriacao { get; set; }
        public DateTime DataAlteração { get; set; }
        public string CodReserva { get; set; }
        public string Paciente { get; set; }
        public DateTime? DataCotacao { get; set; }
        public DateTime? DataCirurgiaAnt { get; set; }
        public DateTime? DataNovaCirurgia { get; set; }
        public int NumCotacao { get; set; }
        public int Status { get; set; }
        public string NumPedido { get; set; }
        public int StatusPedido { get; set; }
    }
}
