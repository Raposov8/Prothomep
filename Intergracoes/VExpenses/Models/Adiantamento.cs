using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses.Models
{
    public class Adiantamento
    {
        public string description { get; set; }
        public string advance_user_id { get; set; }
        public DateTime advance_date {get;set;}
        public decimal value { get; set; }
        public string currency_iso { get; set; }
        public string creator_user_id { get; set; }
    }
}
