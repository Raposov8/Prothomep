using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Diretoria;

namespace SGID.Pages.Relatorios.Diretoria
{
    public class TitulosVencidosInterModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }

        private TOTVSINTERContext Protheus { get; set; }

        public List<TitulosVencidos> Relatorio { get; set; } = new List<TitulosVencidos>();

        public List<string> Clientes { get; set; } = new List<string>();

        public TitulosVencidosInterModel(ApplicationDbContext sgid,TOTVSINTERContext inter)
        {
            SGID = sgid;
            Protheus = inter;
        }

        public void OnGet()
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                DateTime data = DateTime.Now;
                string DataInicio = $"{data.Year}{data.Month.ToString("D2")}{data.Day.ToString("D2")}";

                Relatorio = (from SE10 in Protheus.Se1010s
                             join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                             where SE10.DELET != "*"
                             && (int)(object)SE10.E1Vencrea < (int)(object)DataInicio
                             && SA10.A1Clinter != "G" && SA10.A1Msblql != "1" && SE10.E1Baixa == ""
                             select new TitulosVencidos
                             {
                                 CodCliente = SE10.E1Cliente,
                                 Cliente = SA10.A1Nome,
                                 Valor = SE10.E1Valor,
                                 DataEmissao = $"{SE10.E1Emissao.Substring(6, 2)}/{SE10.E1Emissao.Substring(4, 2)}/{SE10.E1Emissao.Substring(0, 4)}",
                                 DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}"
                             }
                                   )
                                   .OrderByDescending(x => x.Valor).ToList();

                Clientes = (from SE10 in Protheus.Se1010s
                            join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                            where SE10.DELET != "*"
                            && (int)(object)SE10.E1Vencrea < (int)(object)DataInicio
                            && SA10.A1Clinter != "G" && SA10.A1Msblql != "1" && SE10.E1Baixa == ""
                            select new TitulosVencidos
                            {
                                CodCliente = SE10.E1Cliente,
                                Cliente = SA10.A1Nome,
                                Valor = SE10.E1Valor,
                                DataEmissao = $"{SE10.E1Emissao.Substring(6, 2)}/{SE10.E1Emissao.Substring(4, 2)}/{SE10.E1Emissao.Substring(0, 4)}",
                                DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}"
                            }).OrderBy(x => x.Cliente).Select(x => x.Cliente).Distinct().ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "TitulosVencidosInter", user);
            }
        }

        public IActionResult OnPost(string NReduz)
        {
            try
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                DateTime data = DateTime.Now;
                string DataInicio = $"{data.Year}{data.Month.ToString("D2")}{data.Day.ToString("D2")}";

                var query = (from SE10 in Protheus.Se1010s
                             join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                             where SE10.DELET != "*"
                             && (int)(object)SE10.E1Vencrea < (int)(object)DataInicio
                             && SA10.A1Clinter != "G" && SA10.A1Msblql != "1" && SE10.E1Baixa == ""
                             select new TitulosVencidos
                             {
                                 CodCliente = SE10.E1Cliente,
                                 Cliente = SA10.A1Nome,
                                 Valor = SE10.E1Valor,
                                 DataEmissao = $"{SE10.E1Emissao.Substring(6, 2)}/{SE10.E1Emissao.Substring(4, 2)}/{SE10.E1Emissao.Substring(0, 4)}",
                                 DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}"
                             });

                if (NReduz != null)
                {
                    query = query.Where(x => x.Cliente == NReduz);
                }

                Relatorio = query.OrderByDescending(x => x.Valor).ToList();

                Clientes = (from SE10 in Protheus.Se1010s
                            join SA10 in Protheus.Sa1010s on SE10.E1Cliente equals SA10.A1Cod
                            where SE10.DELET != "*"
                            && (int)(object)SE10.E1Vencrea < (int)(object)DataInicio
                            && SA10.A1Clinter != "G" && SA10.A1Msblql != "1" && SE10.E1Baixa == ""
                            select new TitulosVencidos
                            {
                                CodCliente = SE10.E1Cliente,
                                Cliente = SA10.A1Nome,
                                Valor = SE10.E1Valor,
                                DataEmissao = $"{SE10.E1Emissao.Substring(6, 2)}/{SE10.E1Emissao.Substring(4, 2)}/{SE10.E1Emissao.Substring(0, 4)}",
                                DataVencimento = $"{SE10.E1Vencrea.Substring(6, 2)}/{SE10.E1Vencrea.Substring(4, 2)}/{SE10.E1Vencrea.Substring(0, 4)}"
                            }).OrderBy(x => x.Cliente).Select(x => x.Cliente).Distinct().ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "TitulosVencidos Filter", user);
            }

            return Page();
        }
    }
}
