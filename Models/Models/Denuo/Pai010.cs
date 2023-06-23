using System;
using System.Collections.Generic;

namespace SGID.Models.Denuo
{
    public partial class Pai010
    {
        public string PaiFilial { get; set; } = null!;
        public string PaiGrpprc { get; set; } = null!;
        public string PaiDescr { get; set; } = null!;
        public double PaiValor { get; set; }
        public string PaiMsblql { get; set; } = null!;
        public string DELET { get; set; } = null!;
        public int RECNO { get; set; }
    }
}
