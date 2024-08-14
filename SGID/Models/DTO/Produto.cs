using System.Globalization;
namespace SGID.Models.DTO
{
    public class Produto
    {
        public string? Item { get; set; }
        public string? Licit { get; set; }
        public string? Produtos { get; set; }
        public string? Tuss { get; set; }
        public string? Anvisa { get; set; }
        public string? Validade { get; set; }
        public string? Marca { get; set; }
        public double Und { get; set; }
        public string Lote { get; set; }
        public double PrcUnid { get; set; }
        public string? SegUnd { get; set; }
        public double Descon { get; set; }
        public double VlrTotal { get; set; }
        public string? TipoOp { get; set; }
        public string? Check { get; set; }
        public List<string> Lotes { get; set; }
    }
}
