using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses.Models
{
    public class FluxoAprovacao
    {
        public int id { get; set; }
        public int company_id { get; set; }
        public string description { get; set; }
        public string external_id { get; set; }
        public List<Passos> steps { get; set; }
    }
}
