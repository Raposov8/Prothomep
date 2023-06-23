using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class RejeicaoMotivos
    {
        [Key]
        public int Id { get; set; }

        public string Motivo { get; set; }
    }
}
