using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Intergracoes.VExpenses.Models
{
    public class Passos
    {
        public string Operator { get; set; }
        public int entrance_value { get; set; }
        public int order { get; set; }
        public List<Grupos> groups { get; set; }
    }
}
