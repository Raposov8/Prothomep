using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using SGID.Models.Inter;
using SGID.Data;
using SGID.Data.Models;

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
            var campo = "";
            List<Sb1010> Linhas = new List<Sb1010>();
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

                

                int i = 1;

                using (var reader = new StreamReader($"{Pasta}/TemporarioProdutoInter.csv"))
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
                            linha.B1Xsbfami = TesteString.Testestring(values[18], 2);
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
                            linha.B1Xcodanv = "";
                            linha.B1Qe = 0.0;
                            linha.B1Prv1 = 0.0;
                            linha.B1Emin = 0.0;
                            linha.B1Uprc = 0.0;
                            linha.B1Estseg = 0.0;
                            linha.B1Toler = 0.0;
                            linha.B1Qb = 0.0;
                            linha.B1Comis = 0.0;
                            linha.B1Perinv = 0.0;
                            linha.B1Notamin = 0.0;
                            linha.B1Numcop = 0.0;
                            linha.B1Vlrefus = 0.0;
                            linha.B1Numcqpr = 0.0;
                            linha.B1Contcqp = 0.0;
                            linha.B1Qtmidia = 0.0;
                            linha.B1Faixas = 0.0;
                            linha.B1Nropag = 0.0;
                            linha.B1Emax = 0.0;
                            linha.B1Lotven = 0.0;
                            linha.B1Pesbru = 0.0;
                            linha.B1Fracper = 0.0;
                            linha.B1Crdest = 0.0;
                            linha.B1Vlrselo = 0.0;
                            linha.B1Potenci = 0.0;
                            linha.B1Qtdacum = 0.0;
                            linha.B1Pautfet = 0.0;
                            linha.B1Prfdsul = 0.0;
                            linha.B1Codant = "";
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
                            linha.B1Parcei = "";
                            linha.B1Pmicnut = 0.0;
                            linha.B1Pmacnut = 0.0;
                            linha.B1Codqad = "";
                            linha.B1Seloen = "";
                            linha.B1Qbp = 0.0;
                            linha.B1Alfecop = 0.0;
                            linha.B1Alfecst = 0.0;
                            linha.B1Fecop = "";
                            linha.B1Prodrec = "";
                            linha.B1Rprodep = "";
                            linha.B1Vigenc = "";
                            linha.B1Crdpres = 0.0;
                            linha.B1Verean = "";
                            linha.B1Fecp = 0.0;
                            linha.B1Tribmun = "";
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
                            linha.B1Codlan = "";
                            linha.B1Pr43080 = 0.0;
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
                            linha.B1Ajudif = "";
                            linha.B1Meples = "";
                            linha.B1Tpprd = "";
                            linha.B1Base3 = "";
                            linha.B1Desbse3 = "";
                            linha.B1Iat = "";
                            linha.B1Ippt = "";
                            linha.B1Valepre = "";
                            linha.B1Tipobn = "";
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
                            linha.B1Pergart = 0.0;
                            linha.B1Admin = "";
                            linha.B1Porcprl = "";
                            linha.B1Afamad = 0.0;
                            linha.B1Grpti = "";
                            linha.B1Cest = "";
                            linha.B1Mennota = "";
                            linha.B1Xunnego = "";
                            linha.B1Afasemt = 0.0;
                            linha.B1Aimamt = 0.0;
                            linha.B1Afundes = 0.0;
                            linha.B1Impncm = 0.0;
                            linha.B1Conv = 0.0;
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
                            // linha.B1Xcodanv = "";
                            //B1Utpimpx = "";

                            Linhas.Add(linha);

                            Protheus.Sb1010s.Add(linha);
                            Protheus.SaveChanges();
                            i++;

                        }

                    }
                }



                TextoResposta = $"Produtos Importados com Sucesso";
            }
            catch (RankException e)
            {
                Linhas.ForEach(x =>
                {
                    Protheus.Sb1010s.Remove(x);
                    Protheus.SaveChanges();
                });

                TextoResposta = $"Error: Produto:{produto},campo {campo} numero de caracteris excedido";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar Produto Denuo", user);

                return Page();
            }
            catch (FormatException e)
            {
                TextoResposta = $"Error: Checar a formatação do Texto Produto:{produto}";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar Produto Inter", user);

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
                Logger.Log(e, SGID, "Importar Produto Inter", user);

                return Page();
            }
            catch (Exception e)
            {
                TextoResposta = "Error: Contate o TI";

                string user = User.Identity.Name.Split("@")[0].ToUpper();
                Logger.Log(e, SGID, "Importar Produto Inter", user);

                return Page();
            }

            return Page();
        }
    }

}
