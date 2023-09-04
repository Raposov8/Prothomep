using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;
using SGID.Data;
using SGID.Models.Controladoria.FaturamentoNF;
using SGID.Models.Denuo;
using SGID.Models.DTO;
using SGID.Data.Models;
using System;
using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.Common;

namespace SGID.Pages.Importacao
{
    public class ImportarCsvModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        private readonly IWebHostEnvironment _WEB;

        public string TextoResposta { get; set; } = "";

        public ImportarCsvModel(TOTVSDENUOContext denuo,ApplicationDbContext sgid,IWebHostEnvironment web)
        {
            Protheus = denuo;
            SGID = sgid;
            _WEB = web;
        }

        public void OnGet()
        {

        }
        public IActionResult OnPost(IFormCollection Anexos)
        {
            try
            {
                string Pasta = $"{_WEB.WebRootPath}/Temp";

                if (!Directory.Exists(Pasta))
                {
                    Directory.CreateDirectory(Pasta);
                }

                foreach (var anexo in Anexos.Files)
                {
                    string Caminho = $"{Pasta}/Temporario.csv";

                    using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                    {
                        anexo.CopyTo(fileStream);
                    }
                }

                FileInfo file = new FileInfo($"{Pasta}/Temporario.csv");

                var RecnoSw0010 = Protheus.Sw0010s.OrderByDescending(x => x.RECNO).FirstOrDefault().RECNO;
                var RecnoSw1010 = Protheus.Sw1010s.OrderByDescending(x => x.RECNO).FirstOrDefault().RECNO;

                Sw0010 Cabecalho = new Sw0010();
                List<Sw1010> Linhas = new List<Sw1010>();

                int i = 1;
                var Sw0 = "";

                var valorTotal = 0.0;

                using (var reader = new StreamReader($"{Pasta}/Temporario.csv"))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();

                        if (line.Contains(';'))
                        {
                            var values = line.Split(';');

                            if (values[0] == "1")
                            {
                                Cabecalho.RECNO = RecnoSw0010 + 1;

                                var CC = "";
                                if (values[1].Length != 4)
                                {
                                    if (values[1].Length != 3)
                                    {
                                        if (values[1].Length != 2)
                                        {
                                            CC = $"000{values[1]}";
                                        }
                                        else
                                        {
                                            CC = $"00{values[1]}";
                                        }
                                    }
                                    else
                                    {
                                        CC = $"0{values[1]}";
                                    }
                                }

                                Cabecalho.W0Cc = CC;
                                Cabecalho.W0Num = values[2];

                                var data = values[3].Split("/");

                                var datafor = $"{data[2]}";

                                if(datafor.Length != 4)
                                {
                                    datafor = $"20{datafor}";
                                }

                                var dataformat = $"{datafor}{data[1]}{data[0]}";
                                Cabecalho.W0Dt = dataformat;

                                var compra = "";
                                if(values[4].Length != 3)
                                {
                                    if(values[4].Length != 2)
                                    {
                                        compra = $"00{values[4]}";
                                    }
                                    else
                                    {
                                        compra = $"0{values[4]}";
                                    }
                                }

                                Cabecalho.W0Compra = compra;
                                Cabecalho.W0Moeda = values[5];
                                Cabecalho.W0Pole = "01";
                                Cabecalho.DELET = "";
                                Cabecalho.W0C1Num = "";
                                Cabecalho.W0C3Num = "";
                                Cabecalho.W0Claskit = "";
                                Cabecalho.W0Contr = "";
                                Cabecalho.W0DtEmb = "";
                                Cabecalho.W0DtNec = "";
                                Cabecalho.W0Filial = "";
                                Cabecalho.W0Forloj = "";
                                Cabecalho.W0Forn = "";
                                Cabecalho.W0HawbDa = "";
                                Cabecalho.W0Id = "";
                                Cabecalho.W0Kitseri = "";
                                Cabecalho.W0Refer1 = "";
                                Cabecalho.W0Siauto = "";
                                Cabecalho.W0Sikit = "";
                                Cabecalho.W0Solic = "";

                                Sw0 = Cabecalho.W0Num;

                                Protheus.Sw0010s.Add(Cabecalho);

                            }
                            else
                            {

                                var Linha = new Sw1010
                                {
                                    W1Filial = "",
                                    RECNO = RecnoSw1010 + i,
                                    W1CodI = values[1],
                                    W1Fabr = values[2],
                                    W1Forn = values[3],
                                    W1Qtde = Convert.ToDouble(values[4].Replace(",", ".").Replace("$","")),
                                    W1SaldoQ = Convert.ToDouble(values[4].Replace(",", ".").Replace("$", "")),
                                    W1Preco = Convert.ToDouble(values[5].Replace(",", ".").Replace("$", "")),
                                    W1SiNum = Sw0,
                                    W1Fabloj = "01",
                                    W1Forloj = "01",
                                    W1Class = "1",
                                    DELET = "",
                                    W1C3Num = "",
                                    W1Cc = "0001",
                                    W1Codmat = "",
                                    W1Complem = "",
                                    W1Condpg = "",
                                    W1Ctcusto = "",
                                    W1DtCanc = "",
                                    W1DtEmb = "",
                                    W1Dtentr = "",
                                    W1Fluxo = "",
                                    W1Forecas = "",
                                    W1Motcanc = "",
                                    W1NrConc = "",
                                    W1PoNum = "",
                                    W1Posicao = i.ToString("D4"),
                                    W1Posit = "",
                                    W1Segum = "",
                                    W1Status = "",
                                    W1Um = "",
                                };

                                valorTotal += Linha.W1Preco * Linha.W1Qtde;

                                Linhas.Add(Linha);

                                Protheus.Sw1010s.Add(Linha);
                                i++;
                            }
                        }

                    }
                }

                Protheus.SaveChanges();

                TextoResposta = $"Valor Total Da SI: {string.Format("{0:N2}", valorTotal).Replace(",", "").Replace(".", ",")}";
            }
            catch (FormatException e)
            {
                TextoResposta = "Error: Erro de formatação, favor corrigir e tentar novamente";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ImportarCSV Denuo", user);

                return Page();
            }
            catch (Exception e)
            {
                TextoResposta = "Error: Contate o TI";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "ImportarCSV Denuo", user);

                return Page();
            }
            
            return Page();
        }
    }
}
