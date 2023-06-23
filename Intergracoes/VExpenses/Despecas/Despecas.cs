using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses.Despecas
{
    public class Despecas
    {
        public int? id { get; set; }
        public int? user_id { get; set; }
        public int? report_id { get; set; }
        public int? device_id { get; set; }
        public decimal? integration_id { get; set; }
        public string external_id { get; set; }
        public int? expense_type_id { get; set; }
        public int? payment_method_id { get; set; }
        public int? paying_company_id { get; set; }
        public int? route_id { get; set; }
        public string reicept_url { get; set; }
        public DateTime? date { get; set; }
        public decimal? value { get; set; }
        public string title { get; set; }
        public string validate { get; set; }
        public string observation { get; set; }
        public int? rejected { get; set; }
        public bool? on { get; set; }
        public bool? reimbursable { get; set; }
        public decimal? mileage { get; set; }
        public decimal? mileage_value { get; set; }
        public string original_currency_iso { get; set; }
        public decimal? exchange_rate { get; set; }
        public decimal? converted_value { get; set; }
        public string converted_currency_iso { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? updated_at { get; set; }
        public UsuarioRequest? user { get; set; }
        public ReportRequest? report { get; set; }
        public CostCenterRequest? costs_center { get; set; }
        public ExpenseTypeRequest? expense_type { get; set; }
    }
}
