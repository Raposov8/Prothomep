using System;
using System.Collections.Generic;

namespace SGID.Models.Inter;

public partial class Sm2010
{
    public string M2Data { get; set; } = null!;

    public double M2Moeda1 { get; set; }

    public double M2Moeda2 { get; set; }

    public double M2Moeda3 { get; set; }

    public double M2Moeda4 { get; set; }

    public string M2Inform { get; set; } = null!;

    public double M2Moeda5 { get; set; }

    public double M2Txmoed2 { get; set; }

    public double M2Txmoed3 { get; set; }

    public double M2Txmoed4 { get; set; }

    public double M2Txmoed5 { get; set; }

    public string DELET { get; set; } = null!;

    public int RECNO { get; set; }

    public int RECDEL { get; set; }

    public string M2Userlgi { get; set; } = null!;

    public string M2Userlga { get; set; } = null!;

    public string M2Msexp { get; set; } = null!;
}
