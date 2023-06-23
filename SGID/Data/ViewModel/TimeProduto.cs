using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SGID.Data.ViewModel
{
    public class TimeProduto
    {
        [Key]
        public int Id { get; set; }

        [ForeignKey("TimeId")]
        public int TimeId { get; set; }

        public Time Time { get; set; }

        public string Produto { get; set; }
    }
}
