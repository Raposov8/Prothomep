using System;
using System.Collections.Generic;

namespace SGID.Models.Inter;

public partial class Sbf010
{
    public string BfFilial { get; set; } = null!;

    public string BfProduto { get; set; } = null!;

    public string BfLocal { get; set; } = null!;

    public string BfPrior { get; set; } = null!;

    public string BfLocaliz { get; set; } = null!;

    public string BfNumseri { get; set; } = null!;

    public string BfLotectl { get; set; } = null!;

    public string BfNumlote { get; set; } = null!;

    public double BfQuant { get; set; }

    public double BfEmpenho { get; set; }

    public double BfQemppre { get; set; }

    public double BfQtsegum { get; set; }

    public double BfEmpen2 { get; set; }

    public double BfQepre2 { get; set; }

    public string BfDataven { get; set; } = null!;

    public string BfEstfis { get; set; } = null!;

    public string DELET { get; set; } = null!;

    public int RECNO { get; set; }

    public string BfDinvent { get; set; } = null!;
}
