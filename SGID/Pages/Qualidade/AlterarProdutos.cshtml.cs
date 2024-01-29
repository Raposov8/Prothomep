using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using SGID.Data;
using SGID.Data.Models;
using SGID.Models.Denuo;

namespace SGID.Pages.Qualidade
{
    public class AlterarProdutosModel : PageModel
    {
        private TOTVSDENUOContext Protheus { get; set; }
        private ApplicationDbContext SGID { get; set; }

        private readonly IWebHostEnvironment _WEB;

        public string TextoResposta { get; set; } = "";

        public AlterarProdutosModel(TOTVSDENUOContext denuo, ApplicationDbContext sgid, IWebHostEnvironment web)
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
            var campo = "";
            List<Sb1010> Linhas = new List<Sb1010>();
            try
            {
                string Pasta = $"{_WEB.WebRootPath}/Temp";

                var Data = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss").Replace("/", "").Replace(" ", "").Replace(":", "");

                if (!Directory.Exists(Pasta))
                {
                    Directory.CreateDirectory(Pasta);
                }

                foreach (var anexo in Anexos.Files)
                {
                    string Caminho = $"{Pasta}/TemporarioProduto{Data}.csv";

                    using (Stream fileStream = new FileStream(Caminho, FileMode.Create))
                    {
                        anexo.CopyTo(fileStream);
                    }
                }

                var RecnoSB = Protheus.Sb1010s.OrderByDescending(x => x.RECNO).FirstOrDefault().RECNO;

                int i = 1;

                using (var reader = new StreamReader($"{Pasta}/TemporarioProduto{Data}.csv"))
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();

                        if (line.Contains(';'))
                        {
                            var values = line.Split(';');

                            produto = values[1];

                            var linha = new Sb1010
                            {
                                DELET = "",
                                RECNO = RecnoSB + i,
                            };

                            #region CamposTeste
                            campo = "B1_FILIAL";
                            linha.B1Filial = TesteString.Testestring(values[0], 2);
                            campo = "B1_COD";
                            linha.B1Cod = TesteString.Testestring(values[1], 15);
                            campo = "B1_CODUNIT";
                            linha.B1Codunit = TesteString.Testestring(values[2], 15);
                            campo = "B1_DESC";
                            linha.B1Desc = TesteString.Testestring(values[3], 40);
                            campo = "B1_CODITE";
                            linha.B1Codite = TesteString.Testestring(values[4], 27);
                            campo = "B1_DESCING";
                            linha.B1Descing = TesteString.Testestring(values[5], 40);
                            campo = "B1_DESRCOK";
                            linha.B1Desrcok = TesteString.Testestring(values[6], 1);
                            campo = "B1_COMERCI";
                            linha.B1Comerci = TesteString.Testestring(values[7], 1);
                            campo = "B1_TIPO";
                            linha.B1Tipo = TesteString.Testestring(values[8], 2);
                            campo = "B1_LOCPAD";
                            linha.B1Locpad = TesteString.Testestring(values[9], 2);
                            campo = "B1_POSIPI";
                            linha.B1Posipi = TesteString.Testestring(values[10], 10);
                            campo = "B1_ESPECIE";
                            linha.B1Especie = TesteString.Testestring(values[11], 2);
                            campo = "B1_EXNCM";
                            linha.B1ExNcm = TesteString.Testestring(values[12], 3);
                            campo = "B1_EXNBM";
                            linha.B1ExNbm = TesteString.Testestring(values[13], 3);
                            campo = "B1_UM";
                            linha.B1Um = TesteString.Testestring(values[14], 2);
                            campo = "B1_GRUPO";
                            linha.B1Grupo = TesteString.Testestring(values[15], 4);
                            //B1Xsbgrp = values[16];
                            campo = "B1_XFAMILI";
                            linha.B1Xfamili = TesteString.Testestring(values[17], 2);
                            campo = "B1_XSBFAMI";
                            linha.B1Xsbfami = TesteString.Testestring(values[18], 3);
                            campo = "B1_UGRPINT";
                            linha.B1Ugrpint = TesteString.Testestring(values[19], 3);

                            linha.B1Picm = Convert.ToDouble(values[20] == "" ? 0.0 : values[20]);
                            campo = "B1_MSBLQL";
                            linha.B1Msblql = TesteString.Testestring(values[21], 1);
                            campo = "B1_SEGUM";
                            linha.B1Segum = TesteString.Testestring(values[22], 2);

                            linha.B1Ipi = Convert.ToDouble(values[23] == "" ? 0.0 : values[23]);
                            linha.B1Aliqiss = Convert.ToDouble(values[24] == "" ? 0.0 : values[24]);

                            campo = "B1_CODISS";
                            linha.B1Codiss = TesteString.Testestring(values[25], 9);
                            campo = "B1_MOTBLQ";
                            linha.B1Motblq = TesteString.Testestring(values[26], 30);
                            campo = "B1_TE";
                            linha.B1Te = TesteString.Testestring(values[27], 3);
                            campo = "B1_TEIMP";
                            linha.B1Teimp = TesteString.Testestring(values[28], 3);
                            campo = "B1_TEREM";
                            linha.B1Terem = TesteString.Testestring(values[29], 3);
                            campo = "B1_TS";
                            linha.B1Ts = TesteString.Testestring(values[30], 3);
                            campo = "B1_TSREM";
                            linha.B1Tsrem = TesteString.Testestring(values[31], 9);
                            campo = "B1_TSDOA";
                            linha.B1Tsdoa = TesteString.Testestring(values[32], 3);

                            linha.B1Picmret = Convert.ToDouble(values[33] == "" ? 0.0 : values[33]);
                            linha.B1Picment = Convert.ToDouble(values[34] == "" ? 0.0 : values[34]);
                            campo = "B1_ESTFOR";
                            linha.B1Estfor = TesteString.Testestring(values[35], 3);
                            campo = "B1_IMPZFRC";
                            linha.B1Impzfrc = TesteString.Testestring(values[36], 1);
                            campo = "B1_TIPCONV";
                            linha.B1Tipconv = TesteString.Testestring(values[38], 1);
                            campo = "B1_ALTER";
                            linha.B1Alter = TesteString.Testestring(values[39], 15);

                            linha.B1Custd = Convert.ToDouble(values[43] == "" ? 0.0 : values[43]);
                            campo = "B1_UCALSTD";
                            linha.B1Ucalstd = TesteString.Testestring(values[44], 8);
                            campo = "B1_MCUSTD";
                            linha.B1Mcustd = TesteString.Testestring(values[46], 1);

                            linha.B1Peso = Convert.ToDouble(values[47] == "" ? 0.0 : values[47]);
                            campo = "B1_FORPRZ";
                            linha.B1Forprz = TesteString.Testestring(values[49], 3);

                            linha.B1Pe = Convert.ToDouble(values[50] == "" ? 0.0 : values[50]);
                            campo = "B1_TIPE";
                            linha.B1Tipe = TesteString.Testestring(values[51], 1);

                            linha.B1Le = Convert.ToDouble(values[52] == "" ? 0.0 : values[52]);
                            linha.B1Lm = Convert.ToDouble(values[53] == "" ? 0.0 : values[53]);
                            campo = "B1_CONTA";
                            linha.B1Conta = TesteString.Testestring(values[54], 20);
                            campo = "B1_CC";
                            linha.B1Cc = TesteString.Testestring(values[56], 9);
                            campo = "B1_ITEMCC";
                            linha.B1Itemcc = TesteString.Testestring(values[57], 9);
                            campo = "B1_DTREFP1";
                            linha.B1Dtrefp1 = TesteString.Testestring(values[58], 8);
                            campo = "B1_FAMILIA";
                            linha.B1Familia = TesteString.Testestring(values[59], 1);
                            campo = "B1_PROC";
                            linha.B1Proc = TesteString.Testestring(values[60], 6);
                            campo = "B1_LOJPROC";
                            linha.B1Lojproc = TesteString.Testestring(values[61], 2);
                            campo = "B1_APROPRI";
                            linha.B1Apropri = TesteString.Testestring(values[63], 1);
                            campo = "B1_TIPODEC";
                            linha.B1Tipodec = TesteString.Testestring(values[64], 1);
                            campo = "B1_CONTSOC";
                            linha.B1Contsoc = TesteString.Testestring(values[65], 1);
                            campo = "B1_ORIGEM";
                            linha.B1Origem = TesteString.Testestring(values[66], 1);
                            campo = "B1_CODBAR";
                            linha.B1Codbar = TesteString.Testestring(values[67], 15);
                            campo = "B1_GRADE";
                            linha.B1Grade = TesteString.Testestring(values[68], 1);
                            campo = "B1_FORMLOT";
                            linha.B1Formlot = TesteString.Testestring(values[69], 3);
                            campo = "B1_CLASFIS";
                            linha.B1Clasfis = TesteString.Testestring(values[70], 2);
                            campo = "B1_FPCOD";
                            linha.B1Fpcod = TesteString.Testestring(values[71], 2);
                            campo = "B1_FANTASM";
                            linha.B1Fantasm = TesteString.Testestring(values[72], 1);
                            campo = "B1_CONTRAT";
                            linha.B1Contrat = TesteString.Testestring(values[73], 1);
                            campo = "B1_B1RASTRO";
                            linha.B1Rastro = TesteString.Testestring(values[74], 1);
                            campo = "B1_FORAEST";
                            linha.B1Foraest = TesteString.Testestring(values[75], 1);
                            campo = "B1_ANUENTE";
                            linha.B1Anuente = TesteString.Testestring(values[76], 1);
                            campo = "B1_CODOBS";
                            linha.B1Codobs = TesteString.Testestring(values[77], 6);
                            campo = "B1_FABRIC";
                            linha.B1Fabric = TesteString.Testestring(values[79], 20);
                            campo = "B1_GRTRIB";
                            linha.B1Grtrib = TesteString.Testestring(values[81], 3);
                            campo = "B1_MRP";
                            linha.B1Mrp = TesteString.Testestring(values[82], 1);
                            campo = "B1_PRODPAI";
                            linha.B1Prodpai = TesteString.Testestring(values[84], 15);

                            linha.B1Prvalid = Convert.ToDouble(values[85] == "" ? 0.0 : values[85]);
                            campo = "B1_IRRF";
                            linha.B1Irrf = TesteString.Testestring(values[87], 1);
                            campo = "B1_LOCALIZ";
                            linha.B1Localiz = TesteString.Testestring(values[88], 1);
                            campo = "B1_OPERPAD";
                            linha.B1Operpad = TesteString.Testestring(values[89], 2);
                            campo = "B1_IMPORT";
                            linha.B1Import = TesteString.Testestring(values[91], 1);
                            campo = "B1_SITPROD";
                            linha.B1Sitprod = TesteString.Testestring(values[92], 1);
                            campo = "B1_MODELO";
                            linha.B1Modelo = TesteString.Testestring(values[93], 15);
                            campo = "B1_SETOR";
                            linha.B1Setor = TesteString.Testestring(values[94], 2);
                            campo = "B1_BALANCA";
                            linha.B1Balanca = TesteString.Testestring(values[95], 1);
                            campo = "B1_NALNCCA";
                            linha.B1Nalncca = TesteString.Testestring(values[96], 7);
                            campo = "B1_TECLA";
                            linha.B1Tecla = TesteString.Testestring(values[97], 3);
                            campo = "B1_TIPOCQ";
                            linha.B1Tipocq = TesteString.Testestring(values[98], 1);
                            campo = "B1_NALSH";
                            linha.B1Nalsh = TesteString.Testestring(values[99], 8);
                            campo = "B1_SOLICIT";
                            linha.B1Solicit = TesteString.Testestring(values[100], 1);
                            campo = "B1_GRUPCOM";
                            linha.B1Grupcom = TesteString.Testestring(values[101], 6);
                            campo = "B1_AGREGCU";
                            linha.B1Agregcu = TesteString.Testestring(values[102], 1);
                            campo = "B1_DATASUB";
                            linha.B1Datasub = TesteString.Testestring(values[103], 8);
                            campo = "B1_REVATU";
                            linha.B1Revatu = TesteString.Testestring(values[106], 3);
                            campo = "B1_INSS";
                            linha.B1Inss = TesteString.Testestring(values[107], 1);
                            campo = "B1_CODEMB";
                            linha.B1Codemb = TesteString.Testestring(values[108], 20);
                            campo = "B1_ESPECIF";
                            linha.B1Especif = TesteString.Testestring(values[109], 80);
                            campo = "B1_MATPRI";
                            linha.B1MatPri = TesteString.Testestring(values[110], 20);

                            linha.B1Redinss = Convert.ToDouble(values[111] == "" ? 0.0 : values[111]);
                            linha.B1Redirrf = Convert.ToDouble(values[112] == "" ? 0.0 : values[112]);
                            campo = "B1_ALADI";
                            linha.B1Aladi = TesteString.Testestring(values[113], 3);
                            campo = "B1_TABIPI";
                            linha.B1TabIpi = TesteString.Testestring(values[114], 2);
                            campo = "B1_QTDSER";
                            linha.B1Qtdser = TesteString.Testestring(values[115], 1);
                            campo = "B1_GRUDES";
                            linha.B1Grudes = TesteString.Testestring(values[116], 3);

                            linha.B1Redpis = Convert.ToDouble(values[117] == "" ? 0.0 : values[117]);
                            linha.B1Redcof = Convert.ToDouble(values[118] == "" ? 0.0 : values[118]);
                            linha.B1Pcsll = Convert.ToDouble(values[119] == "" ? 0.0 : values[119]);
                            linha.B1Pcofins = Convert.ToDouble(values[120] == "" ? 0.0 : values[120]);
                            linha.B1Ppis = Convert.ToDouble(values[121] == "" ? 0.0 : values[121]);
                            campo = "B1_FLAGSUG";
                            linha.B1Flagsug = TesteString.Testestring(values[122], 1);
                            campo = "B1_CLASSVE";
                            linha.B1Classve = TesteString.Testestring(values[123], 1);
                            campo = "B1_MIDIA";
                            linha.B1Midia = TesteString.Testestring(values[124], 1);

                            linha.B1VlrIpi = Convert.ToDouble(values[126] == "" ? 0.0 : values[126]);
                            campo = "B1_ENVOBR";
                            linha.B1Envobr = TesteString.Testestring(values[127], 1);
                            campo = "B1_SERIE";
                            linha.B1Serie = TesteString.Testestring(values[128], 20);
                            campo = "B1_ISBN";
                            linha.B1Isbn = TesteString.Testestring(values[131], 10);
                            campo = "B1_CORPRI";
                            linha.B1Corpri = TesteString.Testestring(values[132], 6);
                            campo = "B1_CORSEC";
                            linha.B1Corsec = TesteString.Testestring(values[133], 6);
                            campo = "B1_NICONE";
                            linha.B1Nicone = TesteString.Testestring(values[134], 15);
                            campo = "B1_ATRIB1";
                            linha.B1Atrib1 = TesteString.Testestring(values[135], 6);
                            campo = "B1_ATRIB2";
                            linha.B1Atrib2 = TesteString.Testestring(values[136], 6);
                            campo = "B1_ATRIB3";
                            linha.B1Atrib3 = TesteString.Testestring(values[137], 6);
                            campo = "B1_REGSEQ";
                            linha.B1Regseq = TesteString.Testestring(values[138], 6);
                            campo = "B1_TITORIG";
                            linha.B1Titorig = TesteString.Testestring(values[139], 50);
                            campo = "B1_LINGUA";
                            linha.B1Lingua = TesteString.Testestring(values[140], 20);
                            campo = "B1_EDICAO";
                            linha.B1Edicao = TesteString.Testestring(values[141], 3);
                            campo = "B1_OBSISBN";
                            linha.B1Obsisbn = TesteString.Testestring(values[142], 40);
                            campo = "B1_CLVL";
                            linha.B1Clvl = TesteString.Testestring(values[143], 9);
                            campo = "B1_ATIVO";
                            linha.B1Ativo = TesteString.Testestring(values[144], 1);
                            campo = "B1_REQUIS";
                            linha.B1Requis = TesteString.Testestring(values[145], 1);
                            campo = "B1_SELO";
                            linha.B1Selo = TesteString.Testestring(values[147], 1);
                            campo = "B1_OK";
                            linha.B1Ok = TesteString.Testestring(values[149], 4);
                            campo = "B1_USAFEFO";
                            linha.B1Usafefo = TesteString.Testestring(values[150], 1);
                            campo = "B1_CLASSE";
                            linha.B1Classe = TesteString.Testestring(values[151], 6);
                            campo = "B1_TIPCAR";
                            linha.B1Tipcar = TesteString.Testestring(values[153], 6);

                            linha.B1VlrIcm = Convert.ToDouble(values[155] == "" ? 0.0 : values[155]);
                            linha.B1IntIcm = Convert.ToDouble(values[156] == "" ? 0.0 : values[156]);
                            campo = "B1_CNAE";
                            linha.B1Cnae = TesteString.Testestring(values[159], 9);
                            campo = "B1_RETOPER";
                            linha.B1Retoper = TesteString.Testestring(values[160], 1);
                            campo = "B1_FRETISS";
                            linha.B1Fretiss = TesteString.Testestring(values[161], 1);
                            campo = "B1_CODNOR";
                            linha.B1Codnor = TesteString.Testestring(values[162], 3);
                            campo = "B1_CPOTENC";
                            linha.B1Cpotenc = TesteString.Testestring(values[163], 1);
                            campo = "B1_PIS";
                            linha.B1Pis = TesteString.Testestring(values[166], 1);
                            campo = "B1_COFINS";
                            linha.B1Cofins = TesteString.Testestring(values[167], 1);
                            campo = "B1_CSLL";
                            linha.B1Csll = TesteString.Testestring(values[168], 1);
                            campo = "B1_GCCUSTO";
                            linha.B1Gccusto = TesteString.Testestring(values[169], 8);
                            campo = "B1_CCCUSTO";
                            linha.B1Cccusto = TesteString.Testestring(values[170], 9);
                            campo = "B1_CALCFET";
                            linha.B1Calcfet = TesteString.Testestring(values[171], 1);
                            campo = "B1_ESTERIL";
                            linha.B1Esteril = TesteString.Testestring(values[173], 1);
                            campo = "B1_REGANVI";
                            linha.B1Reganvi = TesteString.Testestring(values[174], 15);
                            campo = "B1_DTVALID";
                            linha.B1Dtvalid = TesteString.Testestring(values[175], 8);
                            campo = "B1_XETANVI";
                            linha.B1Xetanvi = TesteString.Testestring(values[176], 3);
                            campo = "B1_XNUMANV";
                            linha.B1Xnumanv = TesteString.Testestring(values[177], 20);
                            campo = "B1_TIPOPAT";
                            linha.B1Tipopat = TesteString.Testestring(values[178], 1);
                            campo = "B1_PATPRIN";
                            linha.B1Patprin = TesteString.Testestring(values[179], 15);
                            campo = "B1_XTPPRD";
                            linha.B1Xtpprd = TesteString.Testestring(values[180], 1);

                            linha.B1Embpadr = Convert.ToDouble(values[181] == "" ? 0.0 : values[181]);
                            campo = "B1_RECESTE";
                            linha.B1Receste = TesteString.Testestring(values[182], 50);
                            campo = "B1_REFEREN";
                            linha.B1Referen = TesteString.Testestring(values[183], 15);
                            campo = "B1_TPREG";
                            linha.B1Tpreg = TesteString.Testestring(values[184], 1);
                            campo = "B1_COMISVE";
                            linha.B1Comisve = TesteString.Testestring(values[185], 1);
                            campo = "B1_COMISGE";
                            linha.B1Comisge = TesteString.Testestring(values[186], 1);

                            linha.B1VlrPis = Convert.ToDouble(values[189] == "" ? 0.0 : values[189]);
                            linha.B1VlrCof = Convert.ToDouble(values[190] == "" ? 0.0 : values[190]);
                            campo = "B1_DESPIMP";
                            linha.B1Despimp = TesteString.Testestring(values[203], 1);
                            campo = "B1_XIMPCOM";
                            linha.B1Ximpcom = TesteString.Testestring(values[204], 1);
                            campo = "B1_FETHAB";
                            linha.B1Fethab = TesteString.Testestring(values[214], 1);
                            campo = "B1_CRICMS";
                            linha.B1Cricms = TesteString.Testestring(values[221], 1);
                            campo = "B1_ESCRIPI";
                            linha.B1Escripi = TesteString.Testestring(values[223], 1);
                            campo = "B1_RICM65";
                            linha.B1Ricm65 = TesteString.Testestring(values[237], 1);
                            campo = "B1_TNATREC";
                            linha.B1Tnatrec = TesteString.Testestring(values[240], 4);
                            campo = "B1_REGESIM";
                            linha.B1Regesim = TesteString.Testestring(values[241], 1);
                            campo = "B1_PRN944I";
                            linha.B1Prn944i = TesteString.Testestring(values[255], 1);
                            campo = "B1_RSATIVO";
                            linha.B1Rsativo = TesteString.Testestring(values[257], 1);
                            campo = "B1_UTPIMP";
                            linha.B1Utpimp = TesteString.Testestring(values[260], 1);
                            campo = "B1_PRODSBP";
                            linha.B1Prodsbp = TesteString.Testestring(values[267], 1);
                            campo = "B1_GARANT";
                            linha.B1Garant = TesteString.Testestring(values[284], 1);
                            campo = "B1_VIGENTE";
                            linha.B1Vigente = TesteString.Testestring(values[287], 13);
                            campo = "B1_GRPCST";
                            linha.B1Grpcst = TesteString.Testestring(values[291], 3);
                            campo = "B1_XFABRIC";
                            linha.B1Xfabric = TesteString.Testestring(values[295], 6);
                            campo = "B1_XSOLIMP";
                            linha.B1Xsolimp = TesteString.Testestring(values[296], 1);
                            campo = "B1_XEXETIQ";
                            linha.B1Xexetiq = TesteString.Testestring(values[297], 1);
                            campo = "B1_CODGTIN";
                            linha.B1Codgtin = TesteString.Testestring(values[303], 15);

                            linha.B1DescP = "";//TesteString.Testestring(values[304],6);
                            linha.B1DescGi = ""; //TesteString.Testestring(values[305],6);
                            linha.B1DescI = "";//TesteString.Testestring(values[306],6);
                            campo = "B1_XIMPCNF";
                            linha.B1Ximpcnf = TesteString.Testestring(values[308], 1);
                            campo = "B1_XTUSS";
                            linha.B1Xtuss = TesteString.Testestring(values[309], 12);
                            //B1Xltrepo = "";
                            //B1Codsec = values[307];
                            //B1Vig = "";
                            #endregion



                            var ProdutoProtheus = Protheus.Sb1010s.FirstOrDefault(x => x.DELET != "*" && x.B1Cod == linha.B1Cod);

                            produto = linha.B1Cod;

                            #region Alteração
                            ProdutoProtheus.B1Codunit = linha.B1Codunit;
                            ProdutoProtheus.B1Desc = linha.B1Desc;
                            ProdutoProtheus.B1Codite = linha.B1Codite;
                            ProdutoProtheus.B1Descing = linha.B1Descing;
                            ProdutoProtheus.B1Desrcok = linha.B1Desrcok;
                            ProdutoProtheus.B1Comerci = linha.B1Comerci;
                            ProdutoProtheus.B1Tipo = linha.B1Tipo;
                            ProdutoProtheus.B1Locpad = linha.B1Locpad;
                            ProdutoProtheus.B1Posipi = linha.B1Posipi;
                            ProdutoProtheus.B1Especie = linha.B1Especie;
                            ProdutoProtheus.B1ExNcm = linha.B1ExNcm;
                            ProdutoProtheus.B1ExNbm = linha.B1ExNbm;
                            ProdutoProtheus.B1Um = linha.B1Um;
                            ProdutoProtheus.B1Grupo = linha.B1Grupo;
                            ProdutoProtheus.B1Xfamili = linha.B1Xfamili;
                            ProdutoProtheus.B1Xsbfami = linha.B1Xsbfami;
                            ProdutoProtheus.B1Ugrpint = linha.B1Ugrpint;
                            ProdutoProtheus.B1Picm = linha.B1Picm;
                            ProdutoProtheus.B1Msblql = linha.B1Msblql;
                            ProdutoProtheus.B1Segum = linha.B1Segum;
                            ProdutoProtheus.B1Ipi = linha.B1Ipi;
                            ProdutoProtheus.B1Aliqiss = linha.B1Aliqiss;
                            ProdutoProtheus.B1Codiss = linha.B1Codiss;
                            ProdutoProtheus.B1Motblq = linha.B1Motblq;
                            ProdutoProtheus.B1Te = linha.B1Te;
                            ProdutoProtheus.B1Teimp = linha.B1Teimp;
                            ProdutoProtheus.B1Terem = linha.B1Terem;
                            ProdutoProtheus.B1Ts = linha.B1Ts;
                            ProdutoProtheus.B1Tsrem = linha.B1Tsrem;
                            ProdutoProtheus.B1Tsdoa = linha.B1Tsdoa;
                            ProdutoProtheus.B1Picmret = linha.B1Picmret;
                            ProdutoProtheus.B1Picment = linha.B1Picment;
                            ProdutoProtheus.B1Estfor = linha.B1Estfor; 
                            ProdutoProtheus.B1Impzfrc = linha.B1Impzfrc;
                            ProdutoProtheus.B1Tipconv = linha.B1Tipconv;
                            ProdutoProtheus.B1Alter = linha.B1Alter;
                            ProdutoProtheus.B1Custd = linha.B1Custd;
                            ProdutoProtheus.B1Ucalstd = linha.B1Ucalstd;
                            ProdutoProtheus.B1Mcustd = linha.B1Mcustd;
                            ProdutoProtheus.B1Peso = linha.B1Peso;
                            ProdutoProtheus.B1Forprz = linha.B1Forprz;
                            ProdutoProtheus.B1Pe = linha.B1Pe;
                            ProdutoProtheus.B1Tipe = linha.B1Tipe;
                            ProdutoProtheus.B1Le = linha.B1Le; 
                            ProdutoProtheus.B1Lm = linha.B1Lm;
                            ProdutoProtheus.B1Conta = linha.B1Conta;
                            ProdutoProtheus.B1Cc = linha.B1Cc;
                            ProdutoProtheus.B1Itemcc = linha.B1Itemcc; 
                            ProdutoProtheus.B1Dtrefp1 = linha.B1Dtrefp1;
                            ProdutoProtheus.B1Familia = linha.B1Familia;
                            ProdutoProtheus.B1Proc = linha.B1Proc;
                            ProdutoProtheus.B1Lojproc = linha.B1Lojproc;
                            ProdutoProtheus.B1Apropri = linha.B1Apropri;
                            ProdutoProtheus.B1Tipodec = linha.B1Tipodec;
                            ProdutoProtheus.B1Contsoc = linha.B1Contsoc;
                            ProdutoProtheus.B1Origem = linha.B1Origem;
                            ProdutoProtheus.B1Codbar = linha.B1Codbar; 
                            ProdutoProtheus.B1Grade = linha.B1Grade;
                            ProdutoProtheus.B1Formlot = linha.B1Formlot;
                            ProdutoProtheus.B1Clasfis = linha.B1Clasfis;
                            ProdutoProtheus.B1Fpcod = linha.B1Fpcod;
                            ProdutoProtheus.B1Fantasm = linha.B1Fantasm; 
                            ProdutoProtheus.B1Contrat = linha.B1Contrat;
                            ProdutoProtheus.B1Rastro = linha.B1Rastro;
                            ProdutoProtheus.B1Foraest = linha.B1Foraest;
                            ProdutoProtheus.B1Anuente = linha.B1Anuente;
                            ProdutoProtheus.B1Codobs = linha.B1Codobs;
                            ProdutoProtheus.B1Fabric = linha.B1Fabric;
                            ProdutoProtheus.B1Grtrib = linha.B1Grtrib;
                            ProdutoProtheus.B1Mrp = linha.B1Mrp;
                            ProdutoProtheus.B1Prodpai = linha.B1Prodpai;
                            ProdutoProtheus.B1Prvalid = linha.B1Prvalid; 
                            ProdutoProtheus.B1Irrf = linha.B1Irrf;
                            ProdutoProtheus.B1Localiz = linha.B1Localiz;
                            ProdutoProtheus.B1Operpad = linha.B1Operpad;
                            ProdutoProtheus.B1Import = linha.B1Import; 
                            ProdutoProtheus.B1Sitprod = linha.B1Sitprod;
                            ProdutoProtheus.B1Modelo = linha.B1Modelo; 
                            ProdutoProtheus.B1Setor = linha.B1Setor;
                            ProdutoProtheus.B1Balanca = linha.B1Balanca;
                            ProdutoProtheus.B1Nalncca = linha.B1Nalncca;
                            ProdutoProtheus.B1Tecla = linha.B1Tecla;
                            ProdutoProtheus.B1Tipocq = linha.B1Tipocq;
                            ProdutoProtheus.B1Nalsh = linha.B1Nalsh;
                            ProdutoProtheus.B1Solicit = linha.B1Solicit;
                            ProdutoProtheus.B1Grupcom = linha.B1Grupcom;
                            ProdutoProtheus.B1Agregcu = linha.B1Agregcu;
                            ProdutoProtheus.B1Datasub = linha.B1Datasub;
                            ProdutoProtheus.B1Revatu = linha.B1Revatu; 
                            ProdutoProtheus.B1Inss = linha.B1Inss;
                            ProdutoProtheus.B1Codemb = linha.B1Codemb; 
                            ProdutoProtheus.B1Especif = linha.B1Especif;
                            ProdutoProtheus.B1MatPri = linha.B1MatPri; 
                            ProdutoProtheus.B1Redinss = linha.B1Redinss;
                            ProdutoProtheus.B1Redirrf = linha.B1Redirrf;
                            ProdutoProtheus.B1Aladi = linha.B1Aladi;
                            ProdutoProtheus.B1TabIpi = linha.B1TabIpi;
                            ProdutoProtheus.B1Qtdser = linha.B1Qtdser;
                            ProdutoProtheus.B1Grudes = linha.B1Grudes;
                            ProdutoProtheus.B1Redpis = linha.B1Redpis; 
                            ProdutoProtheus.B1Redcof = linha.B1Redcof; 
                            ProdutoProtheus.B1Pcsll = linha.B1Pcsll; 
                            ProdutoProtheus.B1Pcofins = linha.B1Pcofins;
                            ProdutoProtheus.B1Ppis = linha.B1Ppis;
                            ProdutoProtheus.B1Flagsug = linha.B1Flagsug;
                            ProdutoProtheus.B1Classve = linha.B1Classve;
                            ProdutoProtheus.B1Midia = linha.B1Midia; 
                            ProdutoProtheus.B1VlrIpi = linha.B1VlrIpi;
                            ProdutoProtheus.B1Envobr = linha.B1Envobr; 
                            ProdutoProtheus.B1Serie = linha.B1Serie;
                            ProdutoProtheus.B1Isbn = linha.B1Isbn;
                            ProdutoProtheus.B1Corpri = linha.B1Corpri;
                            ProdutoProtheus.B1Corsec = linha.B1Corsec;
                            ProdutoProtheus.B1Nicone = linha.B1Nicone;
                            ProdutoProtheus.B1Atrib1 =linha.B1Atrib1;
                            ProdutoProtheus.B1Atrib2 = linha.B1Atrib2;
                            ProdutoProtheus.B1Atrib3 = linha.B1Atrib3;
                            ProdutoProtheus.B1Regseq = linha.B1Regseq;
                            ProdutoProtheus.B1Titorig = linha.B1Titorig;
                            ProdutoProtheus.B1Lingua = linha.B1Lingua;
                            ProdutoProtheus.B1Edicao = linha.B1Edicao;
                            ProdutoProtheus.B1Obsisbn = linha.B1Obsisbn;
                            ProdutoProtheus.B1Clvl = linha.B1Clvl; 
                            ProdutoProtheus.B1Ativo = linha.B1Ativo;
                            ProdutoProtheus.B1Requis = linha.B1Requis;
                            ProdutoProtheus.B1Selo = linha.B1Selo;
                            ProdutoProtheus.B1Ok = linha.B1Ok;
                            ProdutoProtheus.B1Usafefo = linha.B1Usafefo;
                            ProdutoProtheus.B1Classe = linha.B1Classe;
                            ProdutoProtheus.B1Tipcar = linha.B1Tipcar;
                            ProdutoProtheus.B1VlrIcm = linha.B1VlrIcm;
                            ProdutoProtheus.B1IntIcm = linha.B1IntIcm; 
                            ProdutoProtheus.B1Cnae = linha.B1Cnae;
                            ProdutoProtheus.B1Retoper = linha.B1Retoper;
                            ProdutoProtheus.B1Fretiss = linha.B1Fretiss;
                            ProdutoProtheus.B1Codnor = linha.B1Codnor;
                            ProdutoProtheus.B1Cpotenc = linha.B1Cpotenc;
                            ProdutoProtheus.B1Pis = linha.B1Pis;
                            ProdutoProtheus.B1Cofins = linha.B1Cofins; 
                            ProdutoProtheus.B1Csll = linha.B1Csll;
                            ProdutoProtheus.B1Gccusto = linha.B1Gccusto;
                            ProdutoProtheus.B1Cccusto = linha.B1Cccusto;
                            ProdutoProtheus.B1Calcfet = linha.B1Calcfet;
                            ProdutoProtheus.B1Esteril = linha.B1Esteril;
                            ProdutoProtheus.B1Reganvi = linha.B1Reganvi;
                            ProdutoProtheus.B1Dtvalid = linha.B1Dtvalid;
                            ProdutoProtheus.B1Xetanvi = linha.B1Xetanvi;
                            ProdutoProtheus.B1Xnumanv = linha.B1Xnumanv;
                            ProdutoProtheus.B1Tipopat = linha.B1Tipopat;
                            ProdutoProtheus.B1Patprin = linha.B1Patprin;
                            ProdutoProtheus.B1Xtpprd = linha.B1Xtpprd;
                            ProdutoProtheus.B1Embpadr = linha.B1Embpadr;
                            ProdutoProtheus.B1Receste = linha.B1Receste;
                            ProdutoProtheus.B1Referen = linha.B1Referen;
                            ProdutoProtheus.B1Tpreg = linha.B1Tpreg;
                            ProdutoProtheus.B1Comisve = linha.B1Comisve;
                            ProdutoProtheus.B1Comisge = linha.B1Comisge;
                            ProdutoProtheus.B1VlrPis = linha.B1VlrPis;
                            ProdutoProtheus.B1VlrCof = linha.B1VlrCof;
                            ProdutoProtheus.B1Despimp = linha.B1Despimp;
                            ProdutoProtheus.B1Ximpcom = linha.B1Ximpcom;
                            ProdutoProtheus.B1Fethab = linha.B1Fethab;
                            ProdutoProtheus.B1Cricms = linha.B1Cricms;
                            ProdutoProtheus.B1Escripi = linha.B1Escripi;
                            ProdutoProtheus.B1Ricm65 = linha.B1Ricm65;
                            ProdutoProtheus.B1Tnatrec = linha.B1Tnatrec;
                            ProdutoProtheus.B1Regesim = linha.B1Regesim;
                            ProdutoProtheus.B1Prn944i = linha.B1Prn944i;
                            ProdutoProtheus.B1Rsativo = linha.B1Rsativo;
                            ProdutoProtheus.B1Utpimp = linha.B1Utpimp;
                            ProdutoProtheus.B1Prodsbp = linha.B1Prodsbp;
                            ProdutoProtheus.B1Garant = linha.B1Garant;
                            ProdutoProtheus.B1Vigente = linha.B1Vigente;
                            ProdutoProtheus.B1Grpcst = linha.B1Grpcst;
                            ProdutoProtheus.B1Xfabric = linha.B1Xfabric;
                            ProdutoProtheus.B1Xsolimp = linha.B1Xsolimp;
                            ProdutoProtheus.B1Xexetiq = linha.B1Xexetiq;
                            ProdutoProtheus.B1Codgtin = linha.B1Codgtin;
                            ProdutoProtheus.B1Ximpcnf = linha.B1Ximpcnf;
                            ProdutoProtheus.B1Xtuss = linha.B1Xtuss;
                            
                            #endregion

                            Protheus.Sb1010s.Update(ProdutoProtheus);
                            Protheus.SaveChanges();
                            i++;

                        }
                    }
                }

                TextoResposta = $"Produtos Alterados com Sucesso";
            }
            catch (RankException e)
            {
                TextoResposta = $"Error: Produto:{produto},campo {campo} numero de caracteris excedido";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Alterar Produto Denuo", user);

                return Page();
            }
            catch (FormatException e)
            {

                TextoResposta = $"Error: Checar a formatação do Texto Produto:{produto}";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Alterar Produto Denuo", user);

                return Page();
            }
            catch (DbUpdateException e)
            {

                if (e.InnerException != null)
                {
                    TextoResposta = $"Produto: {produto} Error: {e.InnerException.Message}";
                }
                else
                {
                    TextoResposta = $"Produto: {produto} Error: {e.Message}";
                }

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Alterar Produto Denuo", user);

                return Page();
            }
            catch (NullReferenceException e)
            {
                TextoResposta = $"Error: Produto não existe: {produto}";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Alterar Produto Denuo", user);

                return Page();
            }
            catch (Exception e)
            {
                TextoResposta = "Error: Contate o TI";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Alterar Produto Denuo", user);

                return Page();
            }

            return Page();
        }
    }


}
