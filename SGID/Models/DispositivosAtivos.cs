namespace SGID.Models
{
    public class DispositivosAtivos
    {
        public int Id { get; set; }
        public string Nome { get; set; }
        public string Modelo { get; set; }
        public string? Imei { get; set; }
        public string TipoDispositivo { get; set; }
        public string NomeUsuario { get; set; }
    }
}
