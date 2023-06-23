using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses.Models
{
    public class Moedas
    {

        public decimal priority { get; set; }
        public string iso_code { get; set; }
        public string name { get; set; }
        public string symbol { get; set; }
        public string subunit { get; set; }
        public decimal subunit_to_unit { get; set; }
        public bool symbol_first { get; set; }
        public string html_entity { get; set; }
        public string decimal_mark { get; set; }
        public string thousands_separator { get; set; }
        public decimal iso_numeric { get; set; }    
    }
}
