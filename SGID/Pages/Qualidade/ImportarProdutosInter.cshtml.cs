using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Inter;

namespace SGID.Pages.Qualidade
{
    public class ImportarProdutosInterModel : PageModel
    {
        private TOTVSINTERContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        private readonly IWebHostEnvironment _WEB;

        public string TextoResposta { get; set; } = "";

        public ImportarProdutosInterModel(TOTVSINTERContext denuo, ApplicationDbContext sgid, IWebHostEnvironment web)
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

            var produto = "";
            try
            {
                string Pasta = $"{_WEB.WebRootPath}/Temp";

                if (!Directory.Exists(Pasta))
                {
                    Directory.CreateDirectory(Pasta);
                }

                foreach (var anexo in Anexos.Files)
                {
                    string Caminho = $"{Pasta}/TemporarioProdutoInter.csv";

                    using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                    {
                        anexo.CopyTo(fileStream);
                    }
                }

                FileInfo file = new FileInfo($"{Pasta}/TemporarioProdutoInter.csv");

                var RecnoSB = Protheus.Sb1010s.OrderByDescending(x => x.RECNO).FirstOrDefault().RECNO;

                List<Sb1010> Linhas = new List<Sb1010>();

                int i = 1;

                using (var reader = new StreamReader($"{Pasta}/TemporarioProdutoInter.csv"))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();

                        if (line.Contains(";"))
                        {
                            var values = line.Split(';');

                            produto = values[1];

                            var linha = new Sb1010
                            {
                                DELET = "",
                                RECNO = RecnoSB + i,
                            };
                                linha.B1Filial = values[0];
                                linha.B1Cod = values[1];
                                linha.B1Codunit = values[2];
                                linha.B1Desc = values[3];
                                linha.B1Codite = values[4];
                                linha.B1Descing = values[5];
                                linha.B1Desrcok = values[6];
                                linha.B1Comerci = values[7];
                                linha.B1Tipo = values[8];
                                linha.B1Locpad = values[9];
                                linha.B1Posipi = values[10];
                                linha.B1Especie = values[11];
                                linha.B1ExNcm = values[12];
                                linha.B1ExNbm = values[13];
                                linha.B1Um = values[14];
                                linha.B1Grupo = values[15];
                                //B1Xsbgrp = values[16];
                                linha.B1Xfamili = values[17];
                                linha.B1Xsbfami = values[18];
                                linha.B1Ugrpint = values[19];
                                linha.B1Picm = Convert.ToDouble(values[20] == "" ? 0.0 : values[20]);
                                linha.B1Msblql = values[21];
                                linha.B1Segum = values[22];
                                linha.B1Ipi = Convert.ToDouble(values[23] == "" ? 0.0 : values[23]);
                                linha.B1Aliqiss = Convert.ToDouble(values[24] == "" ? 0.0 : values[24]);
                                linha.B1Codiss = values[25];
                                linha.B1Motblq = values[26];
                                linha.B1Te = values[27];
                                linha.B1Teimp = values[28];
                                linha.B1Terem = values[29];
                                linha.B1Ts = values[30];
                                linha.B1Tsrem = values[31];
                                linha.B1Tsdoa = values[32];
                                linha.B1Picmret = Convert.ToDouble(values[33] == "" ? 0.0 : values[33]);
                                linha.B1Picment = Convert.ToDouble(values[34] == "" ? 0.0 : values[34]);
                                linha.B1Estfor = values[35];
                                linha.B1Impzfrc = values[36];
                                linha.B1Conv = 0.0;
                                linha.B1Tipconv = values[38];
                                linha.B1Alter = values[39];
                                linha.B1Qe = 0.0;
                                linha.B1Prv1 = 0.0;
                                linha.B1Emin = 0.0;
                                linha.B1Custd = Convert.ToDouble(values[43] == "" ? 0.0 : values[43]);
                                linha.B1Ucalstd = values[44];
                                linha.B1Uprc = 0.0;
                                linha.B1Mcustd = values[46];
                                linha.B1Peso = Convert.ToDouble(values[47] == "" ? 0.0 : values[47]);
                                linha.B1Estseg = 0.0;
                                linha.B1Forprz = values[49];
                                linha.B1Pe = Convert.ToDouble(values[50] == "" ? 0.0 : values[50]);
                                linha.B1Tipe = values[51];
                                linha.B1Le = Convert.ToDouble(values[52] == "" ? 0.0 : values[52]);
                                linha.B1Lm = Convert.ToDouble(values[53] == "" ? 0.0 : values[53]);
                                linha.B1Conta = values[54];
                                linha.B1Toler = 0.0;
                                linha.B1Cc = values[56];
                                linha.B1Itemcc = values[57];
                                linha.B1Dtrefp1 = values[58];
                                linha.B1Familia = values[59];
                                linha.B1Proc = values[60];
                                linha.B1Lojproc = values[61];
                                linha.B1Qb = 0.0;
                                linha.B1Apropri = values[63];
                                linha.B1Tipodec = values[64];
                                linha.B1Contsoc = values[65];
                                linha.B1Origem = values[66];
                                linha.B1Codbar = values[67];
                                linha.B1Grade = values[68];
                                linha.B1Formlot = values[69];
                                linha.B1Clasfis = values[70];
                                linha.B1Fpcod = values[71];
                                linha.B1Fantasm = values[72];
                                linha.B1Contrat = values[73];
                                linha.B1Rastro = values[74];
                                linha.B1Foraest = values[75];
                                linha.B1Anuente = values[76];
                                linha.B1Codobs = values[77];
                                linha.B1Comis = 0.0;
                                linha.B1Fabric = values[79];
                                linha.B1Perinv = 0.0;
                                linha.B1Grtrib = values[81];
                                linha.B1Mrp = values[82];
                                linha.B1Notamin = 0.0;
                                linha.B1Prodpai = values[84];
                                linha.B1Prvalid = Convert.ToDouble(values[85] == "" ? 0.0 : values[85]);
                                linha.B1Numcop = 0.0;
                                linha.B1Irrf = values[87];
                                linha.B1Localiz = values[88];
                                linha.B1Operpad = values[89];
                                linha.B1Vlrefus = 0.0;
                                linha.B1Import = values[91];
                                linha.B1Sitprod = values[92];
                                linha.B1Modelo = values[93];
                                linha.B1Setor = values[94];
                                linha.B1Balanca = values[95];
                                linha.B1Nalncca = values[96];
                                linha.B1Tecla = values[97];
                                linha.B1Tipocq = values[98];
                                linha.B1Nalsh = values[99];
                                linha.B1Solicit = values[100];
                                linha.B1Grupcom = values[101];
                                linha.B1Agregcu = values[102];
                                linha.B1Datasub = values[103];
                                linha.B1Numcqpr = 0.0;
                                linha.B1Contcqp = 0.0;
                                linha.B1Revatu = values[106];
                                linha.B1Inss = values[107];
                                linha.B1Codemb = values[108];
                                linha.B1Especif = values[109];
                                linha.B1MatPri = values[110];
                                linha.B1Redinss = Convert.ToDouble(values[111] == "" ? 0.0 : values[111]);
                                linha.B1Redirrf = Convert.ToDouble(values[112] == "" ? 0.0 : values[112]);
                                linha.B1Aladi = values[113];
                                linha.B1TabIpi = values[114];
                                linha.B1Qtdser = values[115];
                                linha.B1Grudes = values[116];
                                linha.B1Redpis = Convert.ToDouble(values[117] == "" ? 0.0 : values[117]);
                                linha.B1Redcof = Convert.ToDouble(values[118] == "" ? 0.0 : values[118]);
                                linha.B1Pcsll = Convert.ToDouble(values[119] == "" ? 0.0 : values[119]);
                                linha.B1Pcofins = Convert.ToDouble(values[120] == "" ? 0.0 : values[120]);
                                linha.B1Ppis = Convert.ToDouble(values[121] == "" ? 0.0 : values[121]);
                                linha.B1Flagsug = values[122];
                                linha.B1Classve = values[123];
                                linha.B1Midia = values[124];
                                linha.B1Qtmidia = 0.0;
                                linha.B1VlrIpi = Convert.ToDouble(values[126] == "" ? 0.0 : values[126]);
                                linha.B1Envobr = values[127];
                                linha.B1Serie = values[128];
                                linha.B1Faixas = 0.0;
                                linha.B1Nropag = 0.0;
                                linha.B1Isbn = values[131];
                                linha.B1Corpri = values[132];
                                linha.B1Corsec = values[133];
                                linha.B1Nicone = values[134];
                                linha.B1Atrib1 = values[135];
                                linha.B1Atrib2 = values[136];
                                linha.B1Atrib3 = values[137];
                                linha.B1Regseq = values[138];
                                linha.B1Titorig = values[139];
                                linha.B1Lingua = values[140];
                                linha.B1Edicao = values[141];
                                linha.B1Obsisbn = values[142];
                                linha.B1Clvl = values[143];
                                linha.B1Ativo = values[144];
                                linha.B1Requis = values[145];
                                linha.B1Emax = 0.0;
                                linha.B1Selo = values[147];
                                linha.B1Lotven = 0.0;
                                linha.B1Ok = values[149];
                                linha.B1Usafefo = values[150];
                                linha.B1Classe = values[151];
                                linha.B1Pesbru = 0.0;
                                linha.B1Tipcar = values[153];
                                linha.B1Fracper = 0.0;
                                linha.B1VlrIcm = Convert.ToDouble(values[155] == "" ? 0.0 : values[155]);
                                linha.B1IntIcm = Convert.ToDouble(values[156] == "" ? 0.0 : values[156]);
                                linha.B1Crdest = 0.0;
                                linha.B1Vlrselo = 0.0;
                                linha.B1Cnae = values[159];
                                linha.B1Retoper = values[160];
                                linha.B1Fretiss = values[161];
                                linha.B1Codnor = values[162];
                                linha.B1Cpotenc = values[163];
                                linha.B1Potenci = 0.0;
                                linha.B1Qtdacum = 0.0;
                                linha.B1Pis = values[166];
                                linha.B1Cofins = values[167];
                                linha.B1Csll = values[168];
                                linha.B1Gccusto = values[169];
                                linha.B1Cccusto = values[170];
                                linha.B1Calcfet = values[171];
                                linha.B1Pautfet = 0.0;
                                linha.B1Esteril = values[173];
                                linha.B1Reganvi = values[174];
                                linha.B1Dtvalid = values[175];
                                linha.B1Xetanvi = values[176];
                                linha.B1Xnumanv = values[177];
                                linha.B1Tipopat = values[178];
                                linha.B1Patprin = values[179];
                                linha.B1Xtpprd = values[180];
                                linha.B1Embpadr = Convert.ToDouble(values[181] == "" ? 0.0 : values[181]);
                                linha.B1Receste = values[182];
                                linha.B1Referen = values[183];
                                linha.B1Tpreg = values[184];
                                linha.B1Comisve = values[185];
                                linha.B1Comisge = values[186];
                                linha.B1Prfdsul = 0.0;
                                linha.B1Codant = "";
                                linha.B1VlrPis = Convert.ToDouble(values[189] == "" ? 0.0 : values[189]);
                                linha.B1VlrCof = Convert.ToDouble(values[190] == "" ? 0.0 : values[190]);
                                linha.B1Xcomis1 = 0.0;
                                linha.B1Xcomis2 = 0.0;
                                linha.B1Xcomis3 = 0.0;
                                linha.B1Xcomis4 = 0.0;
                                linha.B1Xpacot1 = 0.0;
                                linha.B1Xpacot2 = 0.0;
                                linha.B1Xpacot3 = 0.0;
                                linha.B1Xpacot4 = 0.0;
                                linha.B1Xconco1 = 0.0;
                                linha.B1Xconco2 = 0.0;
                                linha.B1Xconco3 = 0.0;
                                linha.B1Xconco4 = 0.0;
                                linha.B1Despimp = values[203];
                                linha.B1Ximpcom = values[204];
                                linha.B1Parcei = "";
                                linha.B1Pmicnut = 0.0;
                                linha.B1Pmacnut = 0.0;
                                linha.B1Codqad = "";
                                linha.B1Seloen = "";
                                linha.B1Qbp = 0.0;
                                linha.B1Alfecop = 0.0;
                                linha.B1Alfecst = 0.0;
                                linha.B1Fecop = "";
                                linha.B1Fethab = values[214];
                                linha.B1Prodrec = "";
                                linha.B1Rprodep = "";
                                linha.B1Vigenc = "";
                                linha.B1Crdpres = 0.0;
                                linha.B1Verean = "";
                                linha.B1Fecp = 0.0;
                                linha.B1Cricms = values[221];
                                linha.B1Tribmun = "";
                                linha.B1Escripi = values[223];
                                linha.B1Cfem = "";
                                linha.B1Cfems = "";
                                linha.B1Cfema = 0.0;
                                linha.B1Ivaaju = "";
                                linha.B1Dtcorte = "";
                                linha.B1Afethab = 0.0;
                                linha.B1Afacs = 0.0;
                                linha.B1Afabov = 0.0;
                                linha.B1Tfethab = "";
                                linha.B1Refbas = "";
                                linha.B1Fustf = "";
                                linha.B1Regriss = "";
                                linha.B1Prdori = "";
                                linha.B1Ricm65 = values[237];
                                linha.B1Codlan = "";
                                linha.B1Pr43080 = 0.0;
                                linha.B1Tnatrec = values[240];
                                linha.B1Regesim = values[241];
                                linha.B1Alfumac = 0.0;
                                linha.B1Tpprod = "";
                                linha.B1Fecpba = 0.0;
                                linha.B1Cricmst = "";
                                linha.B1Difcnae = "";
                                linha.B1Dci = "";
                                linha.B1Dcre = "";
                                linha.B1Dcr = "";
                                linha.B1Dcrii = 0.0;
                                linha.B1Coefdcr = 0.0;
                                linha.B1Princmg = 0.0;
                                linha.B1Alfecrn = 0.0;
                                linha.B1Chassi = "";
                                linha.B1Prn944i = values[255];
                                linha.B1Ajudif = "";
                                linha.B1Rsativo = values[257];
                                linha.B1Meples = "";
                                linha.B1Tpprd = "";
                                linha.B1Utpimp = values[260];
                                linha.B1Base3 = "";
                                linha.B1Desbse3 = "";
                                linha.B1Iat = "";
                                linha.B1Ippt = "";
                                linha.B1Valepre = "";
                                linha.B1Tipobn = "";
                                linha.B1Prodsbp = values[267];
                                linha.B1Lotesbp = 0.0;
                                linha.B1Cargae = "";
                                linha.B1Talla = "";
                                linha.B1Gdodif = "";
                                linha.B1Markup = 0.0;
                                linha.B1Vlcif = 0.0;
                                linha.B1Desbse2 = "";
                                linha.B1Color = "";
                                linha.B1Tipvec = "";
                                linha.B1Estrori = "";
                                linha.B1Base = "";
                                linha.B1Idhist = "";
                                linha.B1Pafmd5 = "";
                                linha.B1Sittrib = "";
                                linha.B1Tpdp = "";
                                linha.B1Base2 = "";
                                linha.B1Garant = values[284];
                                linha.B1Pergart = 0.0;
                                linha.B1Admin = "";
                                linha.B1Vigente = values[287];
                                linha.B1Porcprl = "";
                                linha.B1Afamad = 0.0;
                                linha.B1Grpti = "";
                                linha.B1Grpcst = values[291];
                                linha.B1Cest = "";
                                linha.B1Mennota = "";
                                linha.B1Xunnego = "";
                                linha.B1Xfabric = values[295];
                                linha.B1Xsolimp = values[296];
                                linha.B1Xexetiq = values[297];
                                linha.B1Afasemt = 0.0;
                                linha.B1Aimamt = 0.0;
                                linha.B1Afundes = 0.0;
                                linha.B1Impncm = 0.0;
                                //B1Xltrepo = "";
                                linha.B1Codgtin = values[303];
                                linha.B1DescP = "";// values[304];
                                linha.B1DescGi = "";// values[305];
                                linha.B1DescI = "";// values[306];
                                //B1Codsec = values[307];
                                linha.B1Ximpcnf = values[308];
                                linha.B1Xtuss = values[309];
                                //B1Vig = "";
                                linha.B1Apopro = "";
                                linha.B1Bitmap = "";
                                linha.B1Cnatrec = "";
                                linha.B1Codproc = "";
                                linha.B1Conini = "";
                                linha.B1Datref = "";
                                linha.B1Dtfimnt = "";
                                linha.B1Exdtval = "";
                                linha.B1Grpnatr = "";
                                linha.B1Hrexpo = "";
                                linha.B1Impinte = "";
                                linha.B1Integ = "";
                                linha.B1Mono = "";
                                linha.B1Msexp = "";
                                linha.B1Opc = "";
                                linha.B1Quadpro = "";
                                linha.B1Terum = "";
                                linha.B1Ucom = "";
                                linha.B1Urev = "";
                                linha.B1Userlga = "";
                                linha.B1Userlgi = "";
                            //B1Utpimpx = "";

                            Linhas.Add(linha);

                            Protheus.Sb1010s.Add(linha);
                            i++;

                        }

                    }
                }

                Protheus.SaveChanges();

                TextoResposta = $"Produtos Importados com Sucesso";
            }
            catch (FormatException e)
            {
                TextoResposta = $"Error: Checar a formatação do Texto Produto:{produto}";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar SI Inter", user);

                return Page();
            }
            catch (DbUpdateException e)
            {
                if (e.InnerException != null)
                {
                    TextoResposta = $"Error: {e.InnerException.Message}";
                }
                else
                {
                    TextoResposta = $"Error: {e.Message}";
                }

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar SI Inter", user);

                return Page();
            }
            catch (Exception e)
            {
                TextoResposta = "Error: Contate o TI";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar SI Inter", user);

                return Page();
            }

            return Page();
        }
    }
}
