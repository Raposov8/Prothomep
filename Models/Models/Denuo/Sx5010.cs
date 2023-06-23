using System;
using System.Collections.Generic;

namespace SGID.Models.Denuo
{
    public partial class Sx5010
    {
        public string X5Filial { get; set; } = null!;
        public string X5Tabela { get; set; } = null!;
        public string X5Chave { get; set; } = null!;
        public string X5Descri { get; set; } = null!;
        public string X5Descspa { get; set; } = null!;
        public string X5Desceng { get; set; } = null!;
        public string DELET { get; set; } = null!;
        public long RECNO { get; set; }
        public long RECDEL { get; set; }
    }
}
