using System;
using System.Collections.Generic;

namespace SGID.Models.Inter;

public partial class Pah010
{
    public string PahFilial { get; set; } = null!;

    public string PahCodins { get; set; } = null!;

    public string PahNome { get; set; } = null!;

    public string PahRg { get; set; } = null!;

    public string PahCpf { get; set; } = null!;

    public string DELET { get; set; } = null!;

    public int RECNO { get; set; }

    public string PahMsblql { get; set; } = null!;
    public string PahImagem { get; set; }
    public string PahObs { get; set; }
    public string PahTipocontrato { get; set; }
}
