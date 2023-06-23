using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses.Despecas
{
    public class Apontamento
    {
        public int id { get; set; }
        public int integration_id { get; set; }
        public int expense_id { get; set; }
        public int reimbursable_company_id { get; set; }
        public int percentage { get; set; }
        public string description { get; set; }
        public bool on { get; set; }
        public DateTime created_at { get; set; }
        public DateTime updated_at { get; set; }
    }
}
