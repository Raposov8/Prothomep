using System;
using System.Collections.Generic;

namespace SGID.Models.Inter
{
    public partial class Pa3010
    {
        public string Pa3Filial { get; set; } = null!;
        public string Pa3Codigo { get; set; } = null!;
        public string Pa3Produt { get; set; } = null!;
        public string Pa3Lote { get; set; } = null!;
        public string Pa3Sublot { get; set; } = null!;
        public string Pa3Local { get; set; } = null!;
        public string Pa3Dtvali { get; set; } = null!;
        public double Pa3Qtdlot { get; set; }
        public double Pa3Qtdpat { get; set; }
        public string Pa3Conven { get; set; } = null!;
        public string Pa3Xseq { get; set; } = null!;
        public string Pa3Usergi { get; set; } = null!;
        public string Pa3Userga { get; set; } = null!;
        public string DELET { get; set; } = null!;
        public int RECNO { get; set; }
        public string Pa3Dtexcl { get; set; } = null!;
        public string Pa3Locali { get; set; } = null!;
    }
}
