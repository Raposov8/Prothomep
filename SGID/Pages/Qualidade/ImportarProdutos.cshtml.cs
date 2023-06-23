using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding.Validation;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;

namespace SGID.Pages.Qualidade
{
    public class ImportarProdutosModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        private readonly IWebHostEnvironment _WEB;

        public string TextoResposta { get; set; } = "";

        public ImportarProdutosModel(TOTVSDENUOContext denuo, ApplicationDbContext sgid, IWebHostEnvironment web)
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
                    string Caminho = $"{Pasta}/TemporarioProduto.csv";

                    using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                    {
                        anexo.CopyTo(fileStream);
                    }
                }

                FileInfo file = new FileInfo($"{Pasta}/TemporarioProduto.csv");

                var RecnoSB= Protheus.Sb1010s.OrderByDescending(x => x.RECNO).FirstOrDefault().RECNO;

                List<Sb1010> Linhas = new List<Sb1010>();

                int i = 1;

                using (var reader = new StreamReader($"{Pasta}/TemporarioProduto.csv"))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();

                        if (line.Contains(";"))
                        {
                            var values = line.Split(';');

                            var Linha = new Sb1010
                            {
                                DELET = "",
                                RECNO = RecnoSB + i,
                                B1Filial = values[0],
                                B1Cod = values[1],
                                B1Codunit = values[2],
                                B1Desc = values[3],
                                B1Codite = values[4],
                                B1Descing = values[5],
                                B1Desrcok = values[6],
                                B1Comerci = values[7],
                                B1Tipo = values[8],
                                B1Locpad = values[9],
                                B1Posipi = values[10],
                                B1Especie = values[11],
                                B1ExNcm = values[12],
                                B1ExNbm = values[13],
                                B1Um = values[14],
                                B1Grupo = values[15],
                                B1Xsbgrp = values[16],
                                B1Xfamili = values[17],
                                B1Xsbfami = values[18],
                                B1Ugrpint = values[19],
                                B1Picm = Convert.ToDouble(values[20] == "" ? 0.0 : values[20]),
                                B1Msblql = values[21],
                                B1Segum = values[22],
                                B1Ipi = Convert.ToDouble(values[23] == "" ? 0.0 : values[23]),
                                B1Aliqiss = Convert.ToDouble(values[24] == "" ? 0.0 : values[24]),
                                B1Codiss = values[25],
                                B1Motblq = values[26],
                                B1Te = values[27],
                                B1Teimp = values[28],
                                B1Terem = values[29],
                                B1Ts = values[30],
                                B1Tsrem = values[31],
                                B1Tsdoa = values[32],
                                B1Picmret = Convert.ToDouble(values[33] == "" ? 0.0 : values[33]),
                                B1Picment = Convert.ToDouble(values[34] == "" ? 0.0 : values[34]),
                                B1Estfor = values[35],
                                B1Impzfrc = values[36],
                                B1Conv = 0.0,
                                B1Tipconv = values[38],
                                B1Alter = values[39],
                                B1Qe = 0.0,
                                B1Prv1 = 0.0,
                                B1Emin = 0.0,
                                B1Custd = Convert.ToDouble(values[43] == "" ? 0.0 : values[43]),
                                B1Ucalstd = values[44],
                                B1Uprc = 0.0,
                                B1Mcustd = values[46],
                                B1Peso = Convert.ToDouble(values[47] == "" ? 0.0 : values[47]),
                                B1Estseg = 0.0,
                                B1Forprz = values[49],
                                B1Pe = Convert.ToDouble(values[50] == "" ? 0.0 : values[50]),
                                B1Tipe = values[51],
                                B1Le = Convert.ToDouble(values[52] == "" ? 0.0 : values[52]),
                                B1Lm = Convert.ToDouble(values[53] == "" ? 0.0 : values[53]),
                                B1Conta = values[54],
                                B1Toler = 0.0,
                                B1Cc = values[56],
                                B1Itemcc = values[57],
                                B1Dtrefp1 = values[58],
                                B1Familia = values[59],
                                B1Proc = values[60],
                                B1Lojproc = values[61],
                                B1Qb = 0.0,
                                B1Apropri = values[63],
                                B1Tipodec = values[64],
                                B1Contsoc = values[65],
                                B1Origem = values[66],
                                B1Codbar = values[67],
                                B1Grade = values[68],
                                B1Formlot = values[69],
                                B1Clasfis = values[70],
                                B1Fpcod = values[71],
                                B1Fantasm = values[72],
                                B1Contrat = values[73],
                                B1Rastro = values[74],
                                B1Foraest = values[75],
                                B1Anuente = values[76],
                                B1Codobs = values[77],
                                B1Comis = 0.0,
                                B1Fabric = values[79],
                                B1Perinv = 0.0,
                                B1Grtrib = values[81],
                                B1Mrp = values[82],
                                B1Notamin = 0.0,
                                B1Prodpai = values[84],
                                B1Prvalid = Convert.ToDouble(values[85] == "" ? 0.0 : values[85]),
                                B1Numcop = 0.0,
                                B1Irrf = values[87],
                                B1Localiz = values[88],
                                B1Operpad = values[89],
                                B1Vlrefus = 0.0,
                                B1Import = values[91],
                                B1Sitprod = values[92],
                                B1Modelo = values[93],
                                B1Setor = values[94],
                                B1Balanca = values[95],
                                B1Nalncca = values[96],
                                B1Tecla = values[97],
                                B1Tipocq = values[98],
                                B1Nalsh = values[99],
                                B1Solicit = values[100],
                                B1Grupcom = values[101],
                                B1Agregcu = values[102],
                                B1Datasub = values[103],
                                B1Numcqpr = 0.0,
                                B1Contcqp = 0.0,
                                B1Revatu = values[106],
                                B1Inss = values[107],
                                B1Codemb = values[108],
                                B1Especif = values[109],
                                B1MatPri = values[110],
                                B1Redinss = Convert.ToDouble(values[111] == "" ? 0.0 : values[111]),
                                B1Redirrf = Convert.ToDouble(values[112] == "" ? 0.0 : values[112]),
                                B1Aladi = values[113],
                                B1TabIpi = values[114],
                                B1Qtdser = values[115],
                                B1Grudes = values[116],
                                B1Redpis = Convert.ToDouble(values[117] == "" ? 0.0 : values[117]),
                                B1Redcof = Convert.ToDouble(values[118] == "" ? 0.0 : values[118]),
                                B1Pcsll = Convert.ToDouble(values[119] == "" ? 0.0 : values[119]),
                                B1Pcofins = Convert.ToDouble(values[120] == "" ? 0.0 : values[120]),
                                B1Ppis = Convert.ToDouble(values[121] == "" ? 0.0 : values[121]),
                                B1Flagsug = values[122],
                                B1Classve = values[123],
                                B1Midia = values[124],
                                B1Qtmidia = 0.0,
                                B1VlrIpi = Convert.ToDouble(values[126] == "" ? 0.0 : values[126]),
                                B1Envobr = values[127],
                                B1Serie = values[128],
                                B1Faixas = 0.0,
                                B1Nropag = 0.0,
                                B1Isbn = values[131],
                                B1Corpri = values[132],
                                B1Corsec = values[133],
                                B1Nicone = values[134],
                                B1Atrib1 = values[135],
                                B1Atrib2 = values[136],
                                B1Atrib3 = values[137],
                                B1Regseq = values[138],
                                B1Titorig = values[139],
                                B1Lingua = values[140],
                                B1Edicao = values[141],
                                B1Obsisbn = values[142],
                                B1Clvl = values[143],
                                B1Ativo = values[144],
                                B1Requis = values[145],
                                B1Emax = 0.0,
                                B1Selo = values[147],
                                B1Lotven = 0.0,
                                B1Ok = values[149],
                                B1Usafefo = values[150],
                                B1Classe = values[151],
                                B1Pesbru = 0.0,
                                B1Tipcar = values[153],
                                B1Fracper = 0.0,
                                B1VlrIcm = Convert.ToDouble(values[155] == "" ? 0.0 : values[155]),
                                B1IntIcm = Convert.ToDouble(values[156] == "" ? 0.0 : values[156]),
                                B1Crdest = 0.0,
                                B1Vlrselo = 0.0,
                                B1Cnae = values[159],
                                B1Retoper = values[160],
                                B1Fretiss = values[161],
                                B1Codnor = values[162],
                                B1Cpotenc = values[163],
                                B1Potenci = 0.0,
                                B1Qtdacum = 0.0,
                                B1Pis = values[166],
                                B1Cofins = values[167],
                                B1Csll = values[168],
                                B1Gccusto = values[169],
                                B1Cccusto = values[170],
                                B1Calcfet = values[171],
                                B1Pautfet = 0.0,
                                B1Esteril = values[173],
                                B1Reganvi = values[174],
                                B1Dtvalid = values[175],
                                B1Xetanvi = values[176],
                                B1Xnumanv = values[177],
                                B1Tipopat = values[178],
                                B1Patprin = values[179],
                                B1Xtpprd = values[180],
                                B1Embpadr = Convert.ToDouble(values[181] == "" ? 0.0 : values[181]),
                                B1Receste = values[182],
                                B1Referen = values[183],
                                B1Tpreg = values[184],
                                B1Comisve = values[185],
                                B1Comisge = values[186],
                                B1Prfdsul = 0.0,
                                B1Codant = "",
                                B1VlrPis = Convert.ToDouble(values[189] == "" ? 0.0 : values[189]),
                                B1VlrCof = Convert.ToDouble(values[190] == "" ? 0.0: values[190]),
                                B1Xcomis1 = 0.0,
                                B1Xcomis2 = 0.0,
                                B1Xcomis3 = 0.0,
                                B1Xcomis4 = 0.0,
                                B1Xpacot1 = 0.0,
                                B1Xpacot2 = 0.0,
                                B1Xpacot3 = 0.0,
                                B1Xpacot4 = 0.0,
                                B1Xconco1 = 0.0,
                                B1Xconco2 = 0.0,
                                B1Xconco3 = 0.0,
                                B1Xconco4 = 0.0,
                                B1Despimp = values[203],
                                B1Ximpcom = values[204],
                                B1Parcei = "",
                                B1Pmicnut = 0.0,
                                B1Pmacnut = 0.0,
                                B1Codqad = "",
                                B1Seloen = "",
                                B1Qbp = 0.0,
                                B1Alfecop = 0.0,
                                B1Alfecst = 0.0,
                                B1Fecop = "",
                                B1Fethab = values[214],
                                B1Prodrec = "",
                                B1Rprodep = "",
                                B1Vigenc = "",
                                B1Crdpres = 0.0,
                                B1Verean = "",
                                B1Fecp = 0.0,
                                B1Cricms = values[221],
                                B1Tribmun = "",
                                B1Escripi = values[223],
                                B1Cfem = "",
                                B1Cfems = "",
                                B1Cfema = 0.0,
                                B1Ivaaju = "",
                                B1Dtcorte = "",
                                B1Afethab = 0.0,
                                B1Afacs = 0.0,
                                B1Afabov = 0.0,
                                B1Tfethab = "",
                                B1Refbas = "",
                                B1Fustf = "",
                                B1Regriss = "",
                                B1Prdori = "",
                                B1Ricm65 = values[237],
                                B1Codlan = "",
                                B1Pr43080 = 0.0,
                                B1Tnatrec = values[240],
                                B1Regesim = values[241],
                                B1Alfumac = 0.0,
                                B1Tpprod = "",
                                B1Fecpba = 0.0,
                                B1Cricmst = "",
                                B1Difcnae = "",
                                B1Dci = "",
                                B1Dcre = "",
                                B1Dcr = "",
                                B1Dcrii = 0.0,
                                B1Coefdcr = 0.0,
                                B1Princmg = 0.0,
                                B1Alfecrn = 0.0,
                                B1Chassi = "",
                                B1Prn944i = values[255],
                                B1Ajudif = "",
                                B1Rsativo = values[257],
                                B1Meples = "",
                                B1Tpprd = "",
                                B1Utpimp = values[260],
                                B1Base3 = "",
                                B1Desbse3 = "",
                                B1Iat = "",
                                B1Ippt = "",
                                B1Valepre = "",
                                B1Tipobn = "",
                                B1Prodsbp = values[267],
                                B1Lotesbp = 0.0,
                                B1Cargae = "",
                                B1Talla = "",
                                B1Gdodif = "",
                                B1Markup = 0.0,
                                B1Vlcif = 0.0,
                                B1Desbse2 = "",
                                B1Color = "",
                                B1Tipvec = "",
                                B1Estrori = "",
                                B1Base = "",
                                B1Idhist = "",
                                B1Pafmd5 = "",
                                B1Sittrib = "",
                                B1Tpdp = "",
                                B1Base2 = "",
                                B1Garant = values[284],
                                B1Pergart = 0.0,
                                B1Admin = "",
                                B1Vigente = values[287],
                                B1Porcprl = "",
                                B1Afamad = 0.0,
                                B1Grpti = "",
                                B1Grpcst = values[291],
                                B1Cest = "",
                                B1Mennota = "",
                                B1Xunnego = "",
                                B1Xfabric = values[295],
                                B1Xsolimp = values[296],
                                B1Xexetiq = values[297],
                                B1Afasemt = 0.0,
                                B1Aimamt = 0.0,
                                B1Afundes = 0.0,
                                B1Impncm = 0.0,
                                B1Xltrepo = "",
                                B1Codgtin = values[303],
                                B1DescP = "",// values[304],
                                B1DescGi = "",// values[305],
                                B1DescI = "",// values[306],
                                B1Codsec = values[307],
                                B1Ximpcnf = values[308],
                                B1Xtuss = values[309],
                                B1Vig ="",
                                B1Apopro = "",
                                B1Bitmap = "",
                                B1Cnatrec ="",
                                B1Codproc = "",
                                B1Conini="",
                                B1Datref = "",
                                B1Dtfimnt = "",
                                B1Exdtval = "",
                                B1Grpnatr = "",
                                B1Hrexpo = "",
                                B1Impinte = "",
                                B1Integ = "",
                                B1Mono = "",
                                B1Msexp = "",
                                B1Opc = "",
                                B1Quadpro = "",
                                B1Terum = "",
                                B1Ucom = "",
                                B1Urev = "",
                                B1Userlga = "",
                                B1Userlgi = "",
                                B1Utpimpx = "",

                            };

                            Linhas.Add(Linha);

                            Protheus.Sb1010s.Add(Linha);
                            i++;
                            
                        }

                    }
                }

                Protheus.SaveChanges();

                TextoResposta = $"Produtos Importados com Sucesso";
            }
            catch (FormatException e)
            {
                TextoResposta = "Error: Checar a formatação do Texto";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar SI Denuo", user);

                return Page();
            }
            catch (DbUpdateException e)
            {
                if(e.InnerException != null)
                {
                    TextoResposta = $"Error: {e.InnerException.Message}";
                }
                else
                {
                    TextoResposta = $"Error: {e.Message}";
                }

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar SI Denuo", user);

                return Page();
            }
            catch (Exception e)
            {
                TextoResposta = "Error: Contate o TI";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar SI Denuo", user);

                return Page();
            }

            return Page();
        }
    }
}
