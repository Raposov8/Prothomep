using Intergracoes.VExpenses.Despecas;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses.Models
{
    public class Reportar
    {
            public int id { get; set; }
            public string external_id { get; set; }
            public int? user_id { get; set; }
            public int? device_id { get; set; }
            public string description { get; set; }

            public string status { get; set; }
            public int? approval_stage_id { get; set; }
            public int? approval_user_id { get; set; }
            public DateTime? approval_date { get; set; }
            public int? paying_company_id { get; set; }
            public string payment_date { get; set; }
            public int? payment_method_id { get; set; }
            public string observation { get; set; }
            public bool? on { get; set; }
            public string justification { get; set; }
            public string pdf_link { get; set; }
            public string excel_link { get; set; }
            public DateTime? created_at { get; set; }
            public DateTime? updated_at { get; set; }
            public DespecasRequest? expenses { get; set; }
    }
}
