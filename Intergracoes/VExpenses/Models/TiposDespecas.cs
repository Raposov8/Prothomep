using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses.Models
{
    public class TiposDespecas
    {
        public int id { get; set; }
        public string integration_id { get; set; }
        public string description { get; set; }
        public bool? on { get; set; }
    }
}
