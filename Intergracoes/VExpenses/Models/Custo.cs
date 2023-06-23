using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses.Models
{
    public class Custo
    {
        public int id { get; set; }
        public int? integration_id { get; set; }
        public string? name { get; set; }
        public int? company_group_id { get; set; }
        public bool? on { get; set; }
    }
}
