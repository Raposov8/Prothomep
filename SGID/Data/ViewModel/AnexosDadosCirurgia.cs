using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class AnexosDadosCirurgia
    {
        [Key]
        public int Id { get; set; }
        [ForeignKey("DadosCirurgiaId")]
        public int DadosCirurgiaId { get; set; }
        public DadosCirurgia DadosCirurgia { get; set; }
        public string AnexoCam { get; set; }
    }
}
