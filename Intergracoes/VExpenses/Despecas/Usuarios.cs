using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses.Despecas
{
    public class Usuarios
    {
        public int? id { get; set; }
        public string? integration_id { get; set; }
        public string? external_id { get; set; }
        public int? company_id { get; set; }
        public int? role_id { get; set; }
        public int? approval_flow_id { get; set; }
        public int? expense_limit_policy_id { get; set; }
        public string user_type { get; set; }
        public string name { get; set; }
        public string email { get; set; }
        public int? cpf { get; set; }
        public string? phone1 { get; set; }
        public string? phone2 { get; set; }
        public string? birth_date { get; set; }
        public string? bank { get; set; }
        public string? agency { get; set; }
        public string account { get; set; }
        public bool? confirmed { get; set; }
        public bool? active { get; set; }
        public string created_at { get; set; }
        public string updated_at { get; set; }
    }
}
