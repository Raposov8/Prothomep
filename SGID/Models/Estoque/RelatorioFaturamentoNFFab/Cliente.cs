namespace SGID.Models.Estoque.RelatorioFaturamentoNFFab
{
    public class Cliente
    {
        public string Nome { get; set; }
        public List<NotaFiscal> Notas { get; set; } = new List<NotaFiscal>();
    }
}
