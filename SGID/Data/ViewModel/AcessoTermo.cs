using System.ComponentModel.DataAnnotations;

namespace SGID.Data.ViewModel
{
    public class AcessoTermo
    {
        [Key]
        public int Id { get; set; }
        public string Nome { get; set; }
        public string Caminho { get; set; }
        public int AcessoId { get; set; }
        public SolicitacaoAcesso Acesso { get; set; }
    }
}
