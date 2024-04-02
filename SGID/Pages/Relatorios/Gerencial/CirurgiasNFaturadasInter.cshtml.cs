using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SGID.Models.Inter;
using OfficeOpenXml;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Gerencial.RelatorioCirurgiaNFaturada;

namespace SGID.Pages.Relatorios.Gerencial
{
    [Authorize]
    public class CirurgiasNFaturadasInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }
        public List<Operacao> RelatorioContagem { get; set; } = new List<Operacao>();

        public List<CirurgiasN> RelatorioCirurgias { get; set; } = new List<CirurgiasN>();
        public List<Operacao> RelatorioValor { get; set; } = new List<Operacao>();

        public CirurgiasNFaturadasInterModel(TOTVSINTERContext protheus,ApplicationDbContext sgid)
        {
            Protheus = protheus;
            SGID = sgid;
        }
        public void OnGet()
        {
            try
            {
                var utoper = new string[] { "", "A", "B", "C", "D", "E", "O", "P", "R", "X" };
                var Data = DateTime.Now;

                RelatorioContagem = (from SC50 in Protheus.Sc5010s
                                     join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                     join PA10 in Protheus.Pa1010s on SC60.C6Upatrim equals PA10.Pa1Codigo
                                     where SC50.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC60.DELET != "*" && PA10.DELET != "*" && SC50.C5XDtcir != ""
                                     && !utoper.Contains(SC50.C5Utpoper) && (SC50.C5Nota == "" && !Protheus.Sd2010s.Any(x => x.DELET != "*" && x.D2Filial == SC60.C6Filial && x.D2Pedido == SC60.C6Num))
                                     select new Operacao
                                     {
                                         Emissao = SC50.C5Emissao,
                                         Numero = SC50.C5Num,
                                         Nome = SA10.A1Nome,
                                         Paciente = SC50.C5XNmpac,
                                         Medico = SC50.C5XNmmed,
                                         TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO CIRURGIA" : SC50.C5Utpoper == "Q" ? "FAT. QUEBRADO/PERDIDO" : SC50.C5Utpoper == "T" ? "TRIAGEM DE PEDIDOS" : SC50.C5Utpoper == "V" ? "VENDA SUBDISTRIBUIDOR" : "",
                                         DataCirurgia = SC50.C5XDtcir,
                                     }
                             ).ToList();


                RelatorioContagem.ForEach(x =>
                {
                    x.Periodo = (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 30 ? "ATÉ 30 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 30 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 60 ? "ENTRE 31 E 60 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 60 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 90 ? "ENTRE 61 E 90 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 90 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 180 ? "ENTRE 91 E 180 DIAS" : "> 180";
                });

                RelatorioContagem = RelatorioContagem.GroupBy(x => new
                {
                    x.TipoOper,
                    x.Periodo,
                }).Select(x => new Operacao
                {
                    Periodo = x.Key.Periodo,
                    TipoOper = x.Key.TipoOper,
                    Valor = x.Count()
                }).OrderBy(x => x.TipoOper).ToList();


                var query = (from SC50 in Protheus.Sc5010s
                             join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SA10b in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojaent } equals new { Cliente = SA10b.A1Cod, Loja = SA10b.A1Loja }
                             join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                             where SC50.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC60.DELET != "*" && SA10b.DELET != "*" && SC50.C5XDtcir != ""
                             && !utoper.Contains(SC50.C5Utpoper) && (SC50.C5Nota == "" && !Protheus.Sd2010s.Any(x => x.DELET != "*" && x.D2Filial == SC60.C6Filial && x.D2Pedido == SC60.C6Num))
                             select new CirurgiasN
                             {
                                 NomeVend = SC50.C5Nomvend,
                                 Emissao = SC50.C5Emissao,
                                 Numero = SC50.C5Num,
                                 CliFat = SA10.A1Nome,
                                 CliEnt = SA10b.A1Nome,
                                 Paciente = SC50.C5XNmpac,
                                 TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO CIRURGIA" : SC50.C5Utpoper == "Q" ? "FAT. QUEBRADO/PERDIDO" : SC50.C5Utpoper == "T" ? "TRIAGEM DE PEDIDOS" : SC50.C5Utpoper == "V" ? "VENDA SUBDISTRIBUIDOR" : "",
                                 TipoCli = SA10.A1Clinter == "H" ? "HOSPITAL" : SA10.A1Clinter == "M" ? "MEDICO" : SA10.A1Clinter == "I" ? "INSTRUMENTADOR" : SA10.A1Clinter == "N" ? "NORMAL" : SA10.A1Clinter == "C" ? "CONVENIO" : SA10.A1Clinter == "P" ? "PARTICULAR" : SA10.A1Clinter == "S" ? "SUB-DISTRIBUIDOR" : "",
                                 DataCirurgia = SC50.C5XDtcir,
                                 ValTot = SC60.C6Valor,
                                 Descont = SC60.C6Descont
                             }
                             );
                var query2 = query.GroupBy(x => new
                {
                    x.NomeVend,
                    x.Emissao,
                    x.Numero,
                    x.CliFat,
                    x.CliEnt,
                    x.DataCirurgia,
                    x.Paciente,
                    x.TipoOper,
                    x.TipoCli
                });

                RelatorioCirurgias = query2.Select(x => new CirurgiasN
                {
                    NomeVend = x.Key.NomeVend,
                    Emissao = x.Key.Emissao,
                    Numero = x.Key.Numero,
                    CliFat = x.Key.CliFat,
                    CliEnt = x.Key.CliEnt,
                    Paciente = x.Key.Paciente,
                    TipoOper = x.Key.TipoOper,
                    TipoCli = x.Key.TipoCli,
                    DataCirurgia = x.Key.DataCirurgia,
                    ValTot = x.Sum(c => c.ValTot),
                    Descont = x.Sum(c => c.Descont)
                }).OrderByDescending(x => x.DataCirurgia).ToList();


                RelatorioCirurgias.ForEach(x =>
                {
                    x.Dias = (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays;
                    x.Periodo = (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 30 ? "ATÉ 30 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 30 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 60 ? "ENTRE 31 E 60 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 60 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 90 ? "ENTRE 61 E 90 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 90 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 180 ? "ENTRE 91 E 180 DIAS" : "> 180";
                    x.Order = (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 30 ? "1" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 30 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 60 ? "2" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 60 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 90 ? "3" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 90 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 180 ? "4" : "5";
                });


                RelatorioValor = (from SC50 in Protheus.Sc5010s
                                  join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                  join PA10 in Protheus.Pa1010s on SC60.C6Upatrim equals PA10.Pa1Codigo
                                  where SC50.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC60.DELET != "*" && PA10.DELET != "*" && SC50.C5XDtcir != ""
                                  && !utoper.Contains(SC50.C5Utpoper) && (SC50.C5Nota == "" && !Protheus.Sd2010s.Any(x => x.DELET != "*" && x.D2Filial == SC60.C6Filial && x.D2Pedido == SC60.C6Num))
                                  select new Operacao
                                  {
                                      Emissao = SC50.C5Emissao,
                                      Numero = SC50.C5Num,
                                      Nome = SA10.A1Nome,
                                      Paciente = SC50.C5XNmpac,
                                      Medico = SC50.C5XNmmed,
                                      TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO CIRURGIA" : SC50.C5Utpoper == "Q" ? "FAT. QUEBRADO/PERDIDO" : SC50.C5Utpoper == "T" ? "TRIAGEM DE PEDIDOS" : SC50.C5Utpoper == "V" ? "VENDA SUBDISTRIBUIDOR" : "",
                                      UPatrim = SC60.C6Upatrim,
                                      DataCirurgia = SC50.C5XDtcir,
                                      Valor = (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven
                                  }
                              ).GroupBy(x => new
                              {
                                  x.Emissao,
                                  x.DataCirurgia,
                                  x.Numero,
                                  x.Nome,
                                  x.Medico,
                                  x.UPatrim,
                                  x.TipoOper,
                              }).Select(x => new Operacao
                              {
                                  Emissao = x.Key.Emissao,
                                  Numero = x.Key.Numero,
                                  Nome = x.Key.Nome,
                                  Medico = x.Key.Medico,
                                  TipoOper = x.Key.TipoOper,
                                  UPatrim = x.Key.UPatrim,
                                  DataCirurgia = x.Key.DataCirurgia,
                                  Valor = x.Sum(c => c.Valor)
                              }).ToList();

                RelatorioValor.ForEach(x =>
                {
                    x.Periodo = (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 30 ? "ATÉ 30 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 30 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 60 ? "ENTRE 31 E 60 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 60 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 90 ? "ENTRE 61 E 90 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 90 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 180 ? "ENTRE 91 E 180 DIAS" : "> 180";
                });


                RelatorioValor = RelatorioValor.GroupBy(x => new
                {
                    x.TipoOper,
                    x.Periodo,
                }).Select(x => new Operacao
                {
                    Periodo = x.Key.Periodo,
                    TipoOper = x.Key.TipoOper,
                    Valor = x.Sum(c => c.Valor)
                }).OrderBy(x => x.TipoOper).ToList();
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiasNFaturadasInter",user);
            }
        }

        public IActionResult OnPostExport()
        {
            try
            {
                #region Consultas
                var utoper = new string[] { "", "A", "B", "C", "D", "E", "O", "P", "R", "X" };
                var Data = DateTime.Now;

                RelatorioContagem = (from SC50 in Protheus.Sc5010s
                                     join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                     join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                     join PA10 in Protheus.Pa1010s on SC60.C6Upatrim equals PA10.Pa1Codigo
                                     where SC50.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC60.DELET != "*" && PA10.DELET != "*" && SC50.C5XDtcir != ""
                                     && !utoper.Contains(SC50.C5Utpoper) && (SC50.C5Nota == "" && !Protheus.Sd2010s.Any(x => x.DELET != "*" && x.D2Filial == SC60.C6Filial && x.D2Pedido == SC60.C6Num))
                                     select new Operacao
                                     {
                                         Emissao = SC50.C5Emissao,
                                         Numero = SC50.C5Num,
                                         Nome = SA10.A1Nome,
                                         Paciente = SC50.C5XNmpac,
                                         Medico = SC50.C5XNmmed,
                                         TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO CIRURGIA" : SC50.C5Utpoper == "Q" ? "FAT. QUEBRADO/PERDIDO" : SC50.C5Utpoper == "T" ? "TRIAGEM DE PEDIDOS" : SC50.C5Utpoper == "V" ? "VENDA SUBDISTRIBUIDOR" : "",
                                         DataCirurgia = SC50.C5XDtcir,
                                     }
                             ).ToList();


                RelatorioContagem.ForEach(x =>
                {
                    x.Periodo = (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 30 ? "ATÉ 30 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 30 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 60 ? "ENTRE 31 E 60 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 60 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 90 ? "ENTRE 61 E 90 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 90 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 180 ? "ENTRE 91 E 180 DIAS" : "> 180";
                });

                RelatorioContagem = RelatorioContagem.GroupBy(x => new
                {
                    x.TipoOper,
                    x.Periodo,
                }).Select(x => new Operacao
                {
                    Periodo = x.Key.Periodo,
                    TipoOper = x.Key.TipoOper,
                    Valor = x.Count()
                }).OrderBy(x => x.TipoOper).ToList();


                var query = (from SC50 in Protheus.Sc5010s
                             join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                             join SA10b in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojaent } equals new { Cliente = SA10b.A1Cod, Loja = SA10b.A1Loja }
                             join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                             where SC50.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC60.DELET != "*" && SA10b.DELET != "*" && SC50.C5XDtcir != ""
                             && !utoper.Contains(SC50.C5Utpoper) && (SC50.C5Nota == "" && !Protheus.Sd2010s.Any(x => x.DELET != "*" && x.D2Filial == SC60.C6Filial && x.D2Pedido == SC60.C6Num))
                             select new CirurgiasN
                             {
                                 NomeVend = SC50.C5Nomvend,
                                 Emissao = SC50.C5Emissao,
                                 Numero = SC50.C5Num,
                                 CliFat = SA10.A1Nome,
                                 CliEnt = SA10b.A1Nome,
                                 Paciente = SC50.C5XNmpac,
                                 TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO CIRURGIA" : SC50.C5Utpoper == "Q" ? "FAT. QUEBRADO/PERDIDO" : SC50.C5Utpoper == "T" ? "TRIAGEM DE PEDIDOS" : SC50.C5Utpoper == "V" ? "VENDA SUBDISTRIBUIDOR" : "",
                                 TipoCli = SA10.A1Clinter == "H" ? "HOSPITAL" : SA10.A1Clinter == "M" ? "MEDICO" : SA10.A1Clinter == "I" ? "INSTRUMENTADOR" : SA10.A1Clinter == "N" ? "NORMAL" : SA10.A1Clinter == "C" ? "CONVENIO" : SA10.A1Clinter == "P" ? "PARTICULAR" : SA10.A1Clinter == "S" ? "SUB-DISTRIBUIDOR" : "",
                                 DataCirurgia = SC50.C5XDtcir,
                                 ValTot = SC60.C6Valor,
                                 Descont = SC60.C6Descont
                             }
                             );
                var query2 = query.GroupBy(x => new
                {
                    x.NomeVend,
                    x.Emissao,
                    x.Numero,
                    x.CliFat,
                    x.CliEnt,
                    x.DataCirurgia,
                    x.Paciente,
                    x.TipoOper,
                    x.TipoCli
                });

                RelatorioCirurgias = query2.Select(x => new CirurgiasN
                {
                    NomeVend = x.Key.NomeVend,
                    Emissao = x.Key.Emissao,
                    Numero = x.Key.Numero,
                    CliFat = x.Key.CliFat,
                    CliEnt = x.Key.CliEnt,
                    Paciente = x.Key.Paciente,
                    TipoOper = x.Key.TipoOper,
                    TipoCli = x.Key.TipoCli,
                    DataCirurgia = x.Key.DataCirurgia,
                    ValTot = x.Sum(c => c.ValTot),
                    Descont = x.Sum(c => c.Descont)
                }).OrderByDescending(x => x.DataCirurgia).ToList();


                RelatorioCirurgias.ForEach(x =>
                {
                    x.Dias = (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays;
                    x.Periodo = (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 30 ? "ATÉ 30 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 30 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 60 ? "ENTRE 31 E 60 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 60 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 90 ? "ENTRE 61 E 90 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 90 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 180 ? "ENTRE 91 E 180 DIAS" : "> 180";
                    x.Order = (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 30 ? "1" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 30 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 60 ? "2" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 60 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 90 ? "3" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 90 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 180 ? "4" : "5";
                });


                RelatorioValor = (from SC50 in Protheus.Sc5010s
                                  join SA10 in Protheus.Sa1010s on new { Cliente = SC50.C5Cliente, Loja = SC50.C5Lojacli } equals new { Cliente = SA10.A1Cod, Loja = SA10.A1Loja }
                                  join SC60 in Protheus.Sc6010s on new { Filial = SC50.C5Filial, Num = SC50.C5Num } equals new { Filial = SC60.C6Filial, Num = SC60.C6Num }
                                  join PA10 in Protheus.Pa1010s on SC60.C6Upatrim equals PA10.Pa1Codigo
                                  where SC50.DELET != "*" && SA10.DELET != "*" && SA10.A1Msblql != "1" && SC60.DELET != "*" && PA10.DELET != "*" && SC50.C5XDtcir != ""
                                  && !utoper.Contains(SC50.C5Utpoper) && (SC50.C5Nota == "" && !Protheus.Sd2010s.Any(x => x.DELET != "*" && x.D2Filial == SC60.C6Filial && x.D2Pedido == SC60.C6Num))
                                  select new Operacao
                                  {
                                      Emissao = SC50.C5Emissao,
                                      Numero = SC50.C5Num,
                                      Nome = SA10.A1Nome,
                                      Paciente = SC50.C5XNmpac,
                                      Medico = SC50.C5XNmmed,
                                      TipoOper = SC50.C5Utpoper == "F" ? "FATURAMENTO CIRURGIA" : SC50.C5Utpoper == "Q" ? "FAT. QUEBRADO/PERDIDO" : SC50.C5Utpoper == "T" ? "TRIAGEM DE PEDIDOS" : SC50.C5Utpoper == "V" ? "VENDA SUBDISTRIBUIDOR" : "",
                                      UPatrim = SC60.C6Upatrim,
                                      DataCirurgia = SC50.C5XDtcir,
                                      Valor = (SC60.C6Qtdven - SC60.C6Qtdent) * SC60.C6Prcven
                                  }
                              ).GroupBy(x => new
                              {
                                  x.Emissao,
                                  x.DataCirurgia,
                                  x.Numero,
                                  x.Nome,
                                  x.Medico,
                                  x.UPatrim,
                                  x.TipoOper,
                              }).Select(x => new Operacao
                              {
                                  Emissao = x.Key.Emissao,
                                  Numero = x.Key.Numero,
                                  Nome = x.Key.Nome,
                                  Medico = x.Key.Medico,
                                  TipoOper = x.Key.TipoOper,
                                  UPatrim = x.Key.UPatrim,
                                  DataCirurgia = x.Key.DataCirurgia,
                                  Valor = x.Sum(c => c.Valor)
                              }).ToList();

                RelatorioValor.ForEach(x =>
                {
                    x.Periodo = (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 30 ? "ATÉ 30 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 30 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 60 ? "ENTRE 31 E 60 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 60 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 90 ? "ENTRE 61 E 90 DIAS" : (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays > 90 && (Data - Convert.ToDateTime($"{x.DataCirurgia.Substring(6, 2)}/{x.DataCirurgia.Substring(4, 2)}/{x.DataCirurgia.Substring(0, 4)}")).TotalDays <= 180 ? "ENTRE 91 E 180 DIAS" : "> 180";
                });


                RelatorioValor = RelatorioValor.GroupBy(x => new
                {
                    x.TipoOper,
                    x.Periodo,
                }).Select(x => new Operacao
                {
                    Periodo = x.Key.Periodo,
                    TipoOper = x.Key.TipoOper,
                    Valor = x.Sum(c => c.Valor)
                }).OrderBy(x => x.TipoOper).ToList();
                #endregion

                using ExcelPackage package = new ExcelPackage();
                package.Workbook.Worksheets.Add("CirurgiasNFaturadas");

                var sheet = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == "CirurgiaNFaturadas");

                sheet.Cells[1, 1].Value = "Tipo de Operação";
                sheet.Cells[1, 2].Value = "Periodo";
                sheet.Cells[1, 3].Value = "Total";

                int i = 2;

                RelatorioContagem.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.TipoOper;
                    sheet.Cells[i, 2].Value = Pedido.Periodo;
                    sheet.Cells[i, 3].Value = Pedido.Valor;
                    i++;
                });

                i++;

                sheet.Cells[i, 1].Value = "Operação";
                sheet.Cells[i, 2].Value = "Cli Fat.";
                sheet.Cells[i, 3].Value = "Vendedor";
                sheet.Cells[i, 4].Value = "DT.Cirurgia";
                sheet.Cells[i, 5].Value = "Emissão";
                sheet.Cells[i, 6].Value = "Num.Pedido";
                sheet.Cells[i, 7].Value = "Cliente Entrega";
                sheet.Cells[i, 8].Value = "Valor";

                i++;

                RelatorioCirurgias.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.TipoOper;
                    sheet.Cells[i, 2].Value = Pedido.CliFat;
                    sheet.Cells[i, 3].Value = Pedido.NomeVend;
                    sheet.Cells[i, 4].Value = $"{Pedido.DataCirurgia.Substring(6, 2)}/{Pedido.DataCirurgia.Substring(4, 2)}/{Pedido.DataCirurgia.Substring(0, 4)}";
                    sheet.Cells[i, 5].Value = $"{Pedido.Emissao.Substring(6, 2)}/{Pedido.Emissao.Substring(4, 2)}/{Pedido.Emissao.Substring(0, 4)}";
                    sheet.Cells[i, 6].Value = Pedido.Numero;
                    sheet.Cells[i, 7].Value = Pedido.CliEnt;
                    sheet.Cells[i, 8].Value = Pedido.ValTot;

                    i++;
                });

                i++;

                sheet.Cells[1, 1].Value = "Tipo de Operação";
                sheet.Cells[1, 2].Value = "Periodo";
                sheet.Cells[1, 3].Value = "Total";

                i++;

                RelatorioValor.ForEach(Pedido =>
                {
                    sheet.Cells[i, 1].Value = Pedido.TipoOper;
                    sheet.Cells[i, 2].Value = Pedido.Periodo;
                    sheet.Cells[i, 3].Value = Pedido.Valor;
                    i++;
                });

                //var format = sheet.Cells[i, 3, i, 4];
                //format.Style.Numberformat.Format = "";

                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                using MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CirurgiaNFaturadasInter.xlsx");
            }
            catch (Exception e)
            {
                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "CirurgiaNFaturadasInter Excel", user);

                return LocalRedirect("/error");
            }
        }
    }
}
