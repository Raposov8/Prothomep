using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;
using SGID.Data.ViewModel;
using SGID.Models.Cirurgias;
using SGID.Models.Denuo;
using SGID.Models.DTO;

namespace SGID.Pages.Cirurgias
{
    [Authorize]
    //(Roles = "Admin,GestorVenda,Venda")
    public class AgendamentosModel : PageModel
    {
        private ApplicationDbContext SGID { get; set; }
		private TOTVSDENUOContext ProtheusDenuo { get; set; }
		private TOTVSINTERContext ProtheusInter { get; set; }

		public List<Agendamentos> CirurgiasSemData { get; set; }
        public AgendamentosModel(ApplicationDbContext sgid, TOTVSDENUOContext protheusDenuo, TOTVSINTERContext protheusInter)
        {
			SGID = sgid;
			ProtheusDenuo = protheusDenuo;
			ProtheusInter = protheusInter;
        }

		public string BaseUrl { get; set; }
		public void OnGet()
        {
			BaseUrl = $"{Request.Scheme}://{Request.Host}";
			CirurgiasSemData = SGID.Agendamentos.Where(x => x.DataCirurgia == null).ToList();

        }

		public JsonResult OnGetEvents()
		{
			try
			{
				var agendamentos = SGID.Agendamentos.Where(x => x.DataCirurgia != null).ToList();

				var events = new List<EventViewModel>();

				agendamentos.ForEach(x =>
				{
                    var cor = x.Procedimento.Contains("ORTOGNATICA") ? "red" : x.Procedimento.Contains("FRATURA") ? "yellow" : x.Procedimento.Contains("IMPLANTE") ? "green" : "blue";


                    events.Add(new EventViewModel()
					{
						Id = x.Id,
						Title = x.Paciente,
						Start = $"{x.DataCirurgia:MM/dd/yyyy HH:mm}",
						AllDay = false,
                        color = cor,
                        textColor = cor == "yellow"?"#20232a": "#ffffff",

                    });
				//color = "yellow"
				});
				return new JsonResult(events.ToArray());
			}
			catch (Exception e)
			{
				string user = User.Identity.Name.Split("@")[0].ToUpper();
				Logger.Log(e, SGID, "Agendamento",user);
			}

			return new JsonResult("");
		}

		public JsonResult OnGetDetails(int id)
        {
			try
			{
				var agendamento = SGID.Agendamentos.Select(x => new
				{
					x.Cliente,
					x.CondPag,
					x.Convenio,
					DataAutorizacao = x.DataAutorizacao.Value.ToString("dd/MM/yyyy"),
					DataCirurgia = x.DataCirurgia.Value.ToString("dd/MM/yyyy HH:mm"),
					x.Hospital,
					x.Id,
				    x.Instrumentador,
					x.Matricula,
					x.Medico,
					x.NumAutorizacao,
					x.Observacao,
					x.Paciente,
					x.Senha,
					x.Tabela,
					x.Tipo,
					x.Vendedor,
					x.Empresa,
					x.Autorizado,
					x.Procedimento
				}).FirstOrDefault(x => x.Id == id);

				var codigos = SGID.ProdutosAgendamentos.Where(x => x.AgendamentoId == agendamento.Id)
					.Select(x => new
					{
						produto = x.CodigoProduto,
						unidade = x.Quantidade,
                        ValorUnidade = x.ValorUnitario,
						valor = x.ValorTotal
					}).ToList();

				var Patrimonios = SGID.PatrimoniosAgendamentos.Where(x => x.AgendamentoId == agendamento.Id).ToList();

				var Avulsos = SGID.AvulsosAgendamento.Where(x => x.AgendamentoId == agendamento.Id).ToList();

				var produtos = new List<Produto>();
				var patrimonios = new List<Patrimonio>();
				var avulsos = new List<Produto>();

                if (agendamento.Empresa == "01")
                {
                    codigos.ForEach(x =>
                    {
                        var produto = ProtheusInter.Sb1010s.Where(c => c.B1Cod == x.produto).Select(c => new Produto
                        {
                            Item = c.B1Cod,
                            Licit = c.B1Solicit,
                            Produtos = c.B1Desc,
                            Tuss = c.B1Xtuss,
                            Anvisa = c.B1Reganvi,
                            Marca = c.B1Fabric,
                            Und = x.unidade,
                            PrcUnid = x.ValorUnidade,
                            SegUnd = c.B1Um,
                            VlrTotal = x.valor,
                            TipoOp = c.B1Tipo,
                        }).First();

                        produtos.Add(produto);
                    });

                    Patrimonios.ForEach(x =>
                    {
                        var view = (from PA10 in ProtheusInter.Pa1010s
                                    join PAC in ProtheusInter.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                    from c in sr.DefaultIfEmpty()
                                    join SA10 in ProtheusInter.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                                    from a in st.DefaultIfEmpty()
                                    where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                                    && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == x.Patrimonio
                                    select new Patrimonio
                                    {
                                        Codigo = x.Codigo,
                                        Descri = PA10.Pa1Despat,
                                        KitBas = PA10.Pa1Kitbas,
                                    }).First();


                        patrimonios.Add(view);
                    });

					Avulsos.ForEach(x =>
					{
                        var avulso = ProtheusInter.Sb1010s.Where(c => c.B1Cod == x.Produto).Select(c => new Produto
                        {
                            Item = c.B1Cod,
                            Licit = c.B1Solicit,
                            Produtos = c.B1Desc,
                            Tuss = c.B1Xtuss,
                            Anvisa = c.B1Reganvi,
                            Marca = c.B1Fabric,
                            Und = x.Quantidade,
                            TipoOp = c.B1Tipo,
                        }).First();

                        avulsos.Add(avulso);

                    });
                }
                else
                {
                    codigos.ForEach(x =>
                    {
                        var produto = ProtheusDenuo.Sb1010s.Where(c => c.B1Cod == x.produto).Select(c => new Produto
                        {
                            Item = c.B1Cod,
                            Licit = c.B1Solicit,
                            Produtos = c.B1Desc,
                            Tuss = c.B1Xtuss,
                            Anvisa = c.B1Reganvi,
                            Marca = c.B1Fabric,
                            Und = x.unidade,
                            PrcUnid = x.ValorUnidade,
                            SegUnd = c.B1Um,
                            VlrTotal = x.valor,
                        }).First();

                        produtos.Add(produto);
                    });

                    Patrimonios.ForEach(x =>
                    {
                        var view = (from PA10 in ProtheusDenuo.Pa1010s
                                    join PAC in ProtheusDenuo.Pac010s on PA10.Pa1Numage equals PAC.PacNumage into sr
                                    from c in sr.DefaultIfEmpty()
                                    join SA10 in ProtheusDenuo.Sa1010s on new { Codigo = c.PacClient, Loja = c.PacLojent } equals new { Codigo = SA10.A1Cod, Loja = SA10.A1Loja } into st
                                    from a in st.DefaultIfEmpty()
                                    where PA10.DELET != "*" && PA10.Pa1Msblql != "1" && PA10.Pa1Status != "B"
                                    && c.DELET != "*" && a.DELET != "*" && PA10.Pa1Despat == x.Patrimonio
                                    select new Patrimonio
                                    {
                                        Codigo = x.Codigo,
                                        Descri = PA10.Pa1Despat,
                                        KitBas = PA10.Pa1Kitbas,
                                    }).First();


                        patrimonios.Add(view);
                    });

                    Avulsos.ForEach(x =>
                    {
                        var avulso = ProtheusDenuo.Sb1010s.Where(c => c.B1Cod == x.Produto).Select(c => new Produto
                        {
                            Item = c.B1Cod,
                            Licit = c.B1Solicit,
                            Produtos = c.B1Desc,
                            Tuss = c.B1Xtuss,
                            Anvisa = c.B1Reganvi,
                            Marca = c.B1Fabric,
                            Und = x.Quantidade,
                            TipoOp = c.B1Tipo,
                        }).First();

                        avulsos.Add(avulso);

                    });
                }

                var agenda = new
				{
					Agendamento = agendamento,
					Produtos = produtos,
					Patris = patrimonios,
                    Avulso = Avulsos
				};

				return new JsonResult(agenda);
			}
			catch (Exception e)
			{
				string user = User.Identity.Name.Split("@")[0].ToUpper();
				Logger.Log(e, SGID, "Agendamento Details",user);
			}

			return new JsonResult("");
		}

		public JsonResult OnPostSetDataCirurgia(int CirurgiaId,DateTime DataCirurgia)
		{
            try
            {
				var agendamento = SGID.Agendamentos.FirstOrDefault(x => x.Id == CirurgiaId);

				agendamento.DataCirurgia = DataCirurgia;

                
                agendamento.DataEntrega = DataCirurgia.AddDays(-1);
                

                SGID.Agendamentos.Update(agendamento);
				SGID.SaveChanges();
				
                return new JsonResult("");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "SetDataCirurgia", user);
            }

            return new JsonResult("");
        }

		public class EventViewModel
		{
			public int Id { get; set; }
			public string Title { get; set; }
			public string Start { get; set; }
			public string End { get; set; }
			public bool AllDay { get; set; }
            public string color { get; set; }
            public string textColor { get; set; }

        }
	}
}

