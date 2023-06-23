using Microsoft.EntityFrameworkCore.Storage.ValueConversion.Internal;

namespace SGID.Models.Patrimonio
{
    public class DISPPA3
    {
        public string Codigo { get; set; }
        public string Produto { get; set; }
        public string Descricao { get; set; }
        public double QtdPat { get; set; }
        public double QtdLot { get; set; }
        public string Lote { get; set; }
    }
}
