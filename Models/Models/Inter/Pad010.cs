using System;
using System.Collections.Generic;

namespace SGID.Models.Inter;

public partial class Pad010
{
    public string PadFilial { get; set; } = null!;

    public string PadItem { get; set; } = null!;

    public string PadNumage { get; set; } = null!;

    public string PadProdut { get; set; } = null!;

    public string PadUm { get; set; } = null!;

    public string PadPatrim { get; set; } = null!;

    public double PadQuant { get; set; }

    public string PadLocal { get; set; } = null!;

    public string PadLote { get; set; } = null!;

    public string PadDtvali { get; set; } = null!;

    public string PadLocali { get; set; } = null!;

    public string PadStatus { get; set; } = null!;

    public string DELET { get; set; } = null!;

    public int RECNO { get; set; }

    public string PadDtempe { get; set; } = null!;

    public string PadHrempe { get; set; } = null!;
}
