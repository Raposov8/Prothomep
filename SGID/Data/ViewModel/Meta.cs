using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SGID.Data.ViewModel
{
    public class Meta
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        public string Mes { get; set; }
        public string Ano { get; set; }
        public double Valor { get; set; }
    }
}
