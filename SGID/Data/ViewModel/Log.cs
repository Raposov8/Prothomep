using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SGID.Data.ViewModel
{
    public class Log
    {
       [Key]
       [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
       public int Id { get; set; }

       [Required]
       public string App { get; set; }

       [Required]
       public string Page { get; set; }

       [Required]
       public string Description { get; set; }

       [Required]
       public DateTime DataError { get; set; }
       
       [Required]
       public string User { get; set; }
    }
}
