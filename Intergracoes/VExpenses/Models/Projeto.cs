using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses.Models
{
    public class Projeto
    {
        public int id { get; set; }
        public string name { get; set; }
        public string company_name { get; set; }
        public string cnpj { get; set; }
        public string address { get; set; }
        public string neighborhood { get; set; }
        public string city { get; set; }
        public string state { get; set; }
        public int zip_code { get; set; }
        public int phone1 { get; set; }
        public int phone2 { get; set; }
        public bool on { get; set; }
    }
}
