using System;
using System.Collections.Generic;

namespace SGID.Models.Denuo
{
    public partial class Pa2010
    {
        public string Pa2Filial { get; set; } = null!;
        public string Pa2Codigo { get; set; } = null!;
        public string Pa2Comp { get; set; } = null!;
        public double Pa2Qtdkit { get; set; }
        public double Pa2Qtdpat { get; set; }
        public string Pa2Xseq { get; set; } = null!;
        public string Pa2Usergi { get; set; } = null!;
        public string Pa2Userga { get; set; } = null!;
        public string Pa2Dtexcl { get; set; } = null!;
        public string DELET { get; set; } = null!;
        public int RECNO { get; set; }
    }
}
