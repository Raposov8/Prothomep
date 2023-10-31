using System;
using System.Collections.Generic;

namespace SGID.Models.Denuo;

public partial class Pag010
{
    public string PagFilial { get; set; } = null!;

    public string PagGrupo { get; set; } = null!;

    public string PagDescr { get; set; } = null!;

    public string DELET { get; set; } = null!;

    public int RECNO { get; set; }
}
