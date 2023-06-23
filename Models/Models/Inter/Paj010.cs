using System;
using System.Collections.Generic;

namespace SGID.Models.Inter
{
    public partial class Paj010
    {
        public string PajFilial { get; set; } = null!;
        public string PajProces { get; set; } = null!;
        public string PajPrccir { get; set; } = null!;
        public string PajCodins { get; set; } = null!;
        public double PajValor { get; set; }
        public string PajStatus { get; set; } = null!;
        public string PajUsualt { get; set; } = null!;
        public string DELET { get; set; } = null!;
        public int RECNO { get; set; }
        public string PajObserv { get; set; } = null!;
    }
}
