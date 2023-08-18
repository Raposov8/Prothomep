using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Data;
using SGID.Models.Denuo;
using SGID.Models.Estoque;

namespace SGID.Pages.Relatorios.Estoque
{
    public class Armazem10Model : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        public List<RelatorioArmazem10> Relatorio { get; set; } = new List<RelatorioArmazem10>();

        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }

        public Armazem10Model(TOTVSDENUOContext denuo,ApplicationDbContext sgid)
        {
            Protheus = denuo;
            SGID = sgid;
        }
        public void OnGet()
        {
        }

        public IActionResult OnPost(DateTime DataInicio,DateTime DataFim)
        {
            Inicio = DataInicio;
            Fim = DataFim;


            Relatorio = (from SD30 in Protheus.Sd3010s
                         join SB10 in Protheus.Sb1010s on SD30.D3Cod equals SB10.B1Cod
                         where SD30.D3Local == "10" && SD30.DELET != "*"
                         && (int)(object)SD30.D3Emissao >= (int)(object)DataInicio.ToString("yyyy/MM/dd").Replace("/", "")
                         && (int)(object)SD30.D3Emissao <= (int)(object)DataFim.ToString("yyyy/MM/dd").Replace("/", "")
                         orderby SD30.D3Emissao
                         select new RelatorioArmazem10
                         {
                             Filial = SD30.D3Filial,
                             CodProd = SD30.D3Cod,
                             DescProd = SB10.B1Desc,
                             DataTransf = $"{SD30.D3Emissao.Substring(6, 2)}/{SD30.D3Emissao.Substring(4, 2)}/{SD30.D3Emissao.Substring(0, 4)}",
                             TipoTransf = SD30.D3Cf == "DE4" ? "ENTRADA" : "SAIDA"
                         }).ToList();



            return Page();
        }
    }
}
