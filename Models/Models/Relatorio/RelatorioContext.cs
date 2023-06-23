using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace SGID.Models.Relatorio
{
    public partial class RelatorioContext : DbContext
    {
        public RelatorioContext()
        {
        }

        public RelatorioContext(DbContextOptions<RelatorioContext> options)
            : base(options)
        {
        }

        public virtual DbSet<AjusteT> AjusteTs { get; set; } = null!;
        public virtual DbSet<Apagar> Apagars { get; set; } = null!;
        public virtual DbSet<Bdgcirval> Bdgcirvals { get; set; } = null!;
        public virtual DbSet<Bdgrlt> Bdgrlts { get; set; } = null!;
        public virtual DbSet<Cirval> Cirvals { get; set; } = null!;
        public virtual DbSet<Clixperm> Clixperms { get; set; } = null!;
        public virtual DbSet<Custointer> Custointers { get; set; } = null!;
        public virtual DbSet<Fatnf> Fatnfs { get; set; } = null!;
        public virtual DbSet<Fatnffab> Fatnffabs { get; set; } = null!;
        public virtual DbSet<Grpconv> Grpconvs { get; set; } = null!;
        public virtual DbSet<Grupo> Grupos { get; set; } = null!;
        public virtual DbSet<IcmsJan22> IcmsJan22s { get; set; } = null!;
        public virtual DbSet<Loteval> Lotevals { get; set; } = null!;
        public virtual DbSet<NcmI> NcmIs { get; set; } = null!;
        public virtual DbSet<NfdemoDenuo> NfdemoDenuos { get; set; } = null!;
        public virtual DbSet<NfdemoInter> NfdemoInters { get; set; } = null!;
        public virtual DbSet<Orcado> Orcados { get; set; } = null!;
        public virtual DbSet<PermClisec> PermClisecs { get; set; } = null!;
        public virtual DbSet<Resumorlt> Resumorlts { get; set; } = null!;
        public virtual DbSet<Resumoval> Resumovals { get; set; } = null!;
        public virtual DbSet<TbLicitIten> TbLicitItens { get; set; } = null!;
        public virtual DbSet<TbLicitProd> TbLicitProds { get; set; } = null!;
        public virtual DbSet<Tempd2rlt> Tempd2rlts { get; set; } = null!;
        public virtual DbSet<Validint> Validints { get; set; } = null!;
        public virtual DbSet<Validint2> Validint2s { get; set; } = null!;
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            modelBuilder.Entity<AjusteT>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("AJUSTE_TS");

                entity.Property(e => e.Cod)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("COD");

                entity.Property(e => e.Ts)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("TS");
            });

            modelBuilder.Entity<Apagar>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("APAGAR");

                entity.Property(e => e.Anomes)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("ANOMES");

                entity.Property(e => e.Codforn)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CODFORN");

                entity.Property(e => e.Descgrupo)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("DESCGRUPO");

                entity.Property(e => e.Descnat)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("DESCNAT");

                entity.Property(e => e.Empresa)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("EMPRESA");

                entity.Property(e => e.Grpnat)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("GRPNAT");

                entity.Property(e => e.Lojaforn)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("LOJAFORN");

                entity.Property(e => e.Natureza)
                    .HasMaxLength(15)
                    .IsUnicode(false)
                    .HasColumnName("NATUREZA");

                entity.Property(e => e.Nomforn)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("NOMFORN");

                entity.Property(e => e.Vlpagar).HasColumnName("VLPAGAR");

                entity.Property(e => e.Vlpago).HasColumnName("VLPAGO");
            });

            modelBuilder.Entity<Bdgcirval>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("BDGCIRVAL");

                entity.Property(e => e.Anomes)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("ANOMES");

                entity.Property(e => e.Budget).HasColumnName("BUDGET");

                entity.Property(e => e.Empresa)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("EMPRESA");

                entity.Property(e => e.Tipo)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("TIPO");
            });

            modelBuilder.Entity<Bdgrlt>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("BDGRLT");

                entity.Property(e => e.Anomes)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("ANOMES");

                entity.Property(e => e.Empresa)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("EMPRESA");

                entity.Property(e => e.Faturado).HasColumnName("FATURADO");

                entity.Property(e => e.Qtdcir).HasColumnName("QTDCIR");

                entity.Property(e => e.Tipo)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("TIPO");

                entity.Property(e => e.Valorizado).HasColumnName("VALORIZADO");
            });

            modelBuilder.Entity<Cirval>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("CIRVAL");

                entity.Property(e => e.Agend)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("AGEND");

                entity.Property(e => e.Cliente)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CLIENTE");

                entity.Property(e => e.Clientrega)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CLIENTREGA");

                entity.Property(e => e.Codgrpconv)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CODGRPCONV");

                entity.Property(e => e.Convenio)
                    .HasMaxLength(40)
                    .IsUnicode(false)
                    .HasColumnName("CONVENIO");

                entity.Property(e => e.Dtcir)
                    .HasMaxLength(8)
                    .IsUnicode(false)
                    .HasColumnName("DTCIR");

                entity.Property(e => e.Emissao)
                    .HasMaxLength(8)
                    .IsUnicode(false)
                    .HasColumnName("EMISSAO");

                entity.Property(e => e.Emissaonf)
                    .HasMaxLength(8)
                    .IsUnicode(false)
                    .HasColumnName("EMISSAONF");

                entity.Property(e => e.Empresa)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("EMPRESA");

                entity.Property(e => e.Filial)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("FILIAL");

                entity.Property(e => e.Grupoconv)
                    .HasMaxLength(55)
                    .IsUnicode(false)
                    .HasColumnName("GRUPOCONV");

                entity.Property(e => e.Loja)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("LOJA");

                entity.Property(e => e.Lojaentrega)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("LOJAENTREGA");

                entity.Property(e => e.Medico)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("MEDICO");

                entity.Property(e => e.Nome)
                    .HasMaxLength(40)
                    .IsUnicode(false)
                    .HasColumnName("NOME");

                entity.Property(e => e.Nomentrega)
                    .HasMaxLength(40)
                    .IsUnicode(false)
                    .HasColumnName("NOMENTREGA");

                entity.Property(e => e.Numnf)
                    .HasMaxLength(9)
                    .IsUnicode(false)
                    .HasColumnName("NUMNF");

                entity.Property(e => e.Paciente)
                    .HasMaxLength(60)
                    .IsUnicode(false)
                    .HasColumnName("PACIENTE");

                entity.Property(e => e.Pedido)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("PEDIDO");

                entity.Property(e => e.Serienf)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("SERIENF");

                entity.Property(e => e.Tipo)
                    .HasMaxLength(16)
                    .IsUnicode(false)
                    .HasColumnName("TIPO");

                entity.Property(e => e.Total).HasColumnName("TOTAL");

                entity.Property(e => e.Totalnf)
                    .HasColumnName("TOTALNF")
                    .HasDefaultValueSql("((0))");

                entity.Property(e => e.Tpbudget)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("TPBUDGET");

                entity.Property(e => e.Vendedor)
                    .HasMaxLength(40)
                    .IsUnicode(false)
                    .HasColumnName("VENDEDOR");
            });

            modelBuilder.Entity<Clixperm>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("CLIXPERM");

                entity.Property(e => e.Codcli)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CODCLI");

                entity.Property(e => e.Codloja)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("CODLOJA");

                entity.Property(e => e.Permanente)
                    .HasMaxLength(40)
                    .IsUnicode(false)
                    .HasColumnName("PERMANENTE");

                entity.Property(e => e.Qtd).HasColumnName("QTD");
            });

            modelBuilder.Entity<Custointer>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("CUSTOINTER");

                entity.Property(e => e.Custo).HasColumnName("CUSTO");

                entity.Property(e => e.Produto)
                    .HasMaxLength(25)
                    .IsUnicode(false)
                    .HasColumnName("PRODUTO");
            });

            modelBuilder.Entity<Fatnf>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("FATNF");

                entity.HasIndex(e => new { e.Empresa, e.Filial, e.Emissao }, "NonClusteredIndex-20170413-165324");

                entity.HasIndex(e => new { e.Tipo, e.Emissao }, "NonClusteredIndex-20170920-084537");

                entity.Property(e => e.Clifor)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CLIFOR");

                entity.Property(e => e.Descon).HasColumnName("DESCON");

                entity.Property(e => e.Dtdigit)
                    .HasMaxLength(8)
                    .IsUnicode(false)
                    .HasColumnName("DTDIGIT");

                entity.Property(e => e.Emissao)
                    .HasMaxLength(8)
                    .IsUnicode(false)
                    .HasColumnName("EMISSAO");

                entity.Property(e => e.Empresa)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("EMPRESA");

                entity.Property(e => e.Est)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("EST");

                entity.Property(e => e.Filial)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("FILIAL");

                entity.Property(e => e.Loja)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("LOJA");

                entity.Property(e => e.Mun)
                    .HasMaxLength(30)
                    .IsUnicode(false)
                    .HasColumnName("MUN");

                entity.Property(e => e.Nf)
                    .HasMaxLength(9)
                    .IsUnicode(false)
                    .HasColumnName("NF");

                entity.Property(e => e.Nome)
                    .HasMaxLength(40)
                    .IsUnicode(false)
                    .HasColumnName("NOME");

                entity.Property(e => e.Serie)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("SERIE");

                entity.Property(e => e.Tipo)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("TIPO");

                entity.Property(e => e.Tipofd)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("TIPOFD");

                entity.Property(e => e.Total).HasColumnName("TOTAL");

                entity.Property(e => e.Totalbrut).HasColumnName("TOTALBRUT");

                entity.Property(e => e.Valicm).HasColumnName("VALICM");

                entity.Property(e => e.Valipi).HasColumnName("VALIPI");
            });

            modelBuilder.Entity<Fatnffab>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("FATNFFAB");

                entity.Property(e => e.Clifor)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CLIFOR");

                entity.Property(e => e.Descon).HasColumnName("DESCON");

                entity.Property(e => e.Dtdigit)
                    .HasMaxLength(8)
                    .IsUnicode(false)
                    .HasColumnName("DTDIGIT");

                entity.Property(e => e.Emissao)
                    .HasMaxLength(8)
                    .IsUnicode(false)
                    .HasColumnName("EMISSAO");

                entity.Property(e => e.Empresa)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("EMPRESA");

                entity.Property(e => e.Est)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("EST");

                entity.Property(e => e.Fabricante)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("FABRICANTE");

                entity.Property(e => e.Filial)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("FILIAL");

                entity.Property(e => e.Loja)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("LOJA");

                entity.Property(e => e.Mun)
                    .HasMaxLength(30)
                    .IsUnicode(false)
                    .HasColumnName("MUN");

                entity.Property(e => e.Nf)
                    .HasMaxLength(9)
                    .IsUnicode(false)
                    .HasColumnName("NF");

                entity.Property(e => e.Nome)
                    .HasMaxLength(40)
                    .IsUnicode(false)
                    .HasColumnName("NOME");

                entity.Property(e => e.Serie)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("SERIE");

                entity.Property(e => e.Tipo)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("TIPO");

                entity.Property(e => e.Tipofd)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("TIPOFD");

                entity.Property(e => e.Total).HasColumnName("TOTAL");

                entity.Property(e => e.Valicm).HasColumnName("VALICM");

                entity.Property(e => e.Valipi).HasColumnName("VALIPI");
            });

            modelBuilder.Entity<Grpconv>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("GRPCONV");

                entity.Property(e => e.Codigo)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CODIGO");

                entity.Property(e => e.Descgrupo)
                    .HasMaxLength(55)
                    .IsUnicode(false)
                    .HasColumnName("DESCGRUPO");
            });

            modelBuilder.Entity<Grupo>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("GRUPOS");

                entity.Property(e => e.Codcli)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CODCLI");

                entity.Property(e => e.Grupo1)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("GRUPO");

                entity.Property(e => e.Loja)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("LOJA");
            });

            modelBuilder.Entity<IcmsJan22>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("ICMS_JAN22");

                entity.Property(e => e.Icm).HasColumnName("ICM");

                entity.Property(e => e.Ncm)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("NCM");

                entity.Property(e => e.Produto)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("PRODUTO");

                entity.Property(e => e.Tesimp)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("TESIMP");

                entity.Property(e => e.Tespad)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("TESPAD");

                entity.Property(e => e.Tipotrib)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("TIPOTRIB");
            });

            modelBuilder.Entity<Loteval>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("LOTEVAL");

                entity.Property(e => e.Lote)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("LOTE");

                entity.Property(e => e.Produto)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("PRODUTO");
            });

            modelBuilder.Entity<NcmI>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("NCM_I");

                entity.Property(e => e.Clasfis)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("CLASFIS");

                entity.Property(e => e.Codigo)
                    .HasMaxLength(30)
                    .IsUnicode(false)
                    .HasColumnName("CODIGO");

                entity.Property(e => e.Ncm)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("NCM");

                entity.Property(e => e.Pcofins).HasColumnName("PCOFINS");

                entity.Property(e => e.Picm).HasColumnName("PICM");

                entity.Property(e => e.Pipi).HasColumnName("PIPI");

                entity.Property(e => e.Ppis).HasColumnName("PPIS");

                entity.Property(e => e.Tes)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("TES");

                entity.Property(e => e.Utpimp)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("UTPIMP");
            });

            modelBuilder.Entity<NfdemoDenuo>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("NFDEMO_DENUO");

                entity.Property(e => e.Cliente)
                    .HasMaxLength(40)
                    .IsUnicode(false)
                    .HasColumnName("CLIENTE");

                entity.Property(e => e.Codcli)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CODCLI");

                entity.Property(e => e.Emissao)
                    .HasColumnType("date")
                    .HasColumnName("EMISSAO");

                entity.Property(e => e.Filial)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("FILIAL");

                entity.Property(e => e.Loja)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("LOJA");

                entity.Property(e => e.Nf)
                    .HasMaxLength(9)
                    .IsUnicode(false)
                    .HasColumnName("NF");

                entity.Property(e => e.Serie)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("SERIE");

                entity.Property(e => e.Tipo)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("TIPO");
            });

            modelBuilder.Entity<NfdemoInter>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("NFDEMO_INTER");

                entity.Property(e => e.Cliente)
                    .HasMaxLength(40)
                    .IsUnicode(false)
                    .HasColumnName("CLIENTE");

                entity.Property(e => e.Codcli)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CODCLI");

                entity.Property(e => e.Emissao)
                    .HasColumnType("date")
                    .HasColumnName("EMISSAO");

                entity.Property(e => e.Filial)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("FILIAL");

                entity.Property(e => e.Loja)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("LOJA");

                entity.Property(e => e.Nf)
                    .HasMaxLength(9)
                    .IsUnicode(false)
                    .HasColumnName("NF");

                entity.Property(e => e.Serie)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("SERIE");

                entity.Property(e => e.Tipo)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("TIPO");
            });

            modelBuilder.Entity<Orcado>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("ORCADO");

                entity.Property(e => e.Anomes)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("ANOMES");

                entity.Property(e => e.Ccusto)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("CCUSTO");

                entity.Property(e => e.Empresa)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("EMPRESA");

                entity.Property(e => e.Fornecedor)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("FORNECEDOR");

                entity.Property(e => e.Gestor)
                    .HasMaxLength(30)
                    .IsUnicode(false)
                    .HasColumnName("GESTOR");

                entity.Property(e => e.Loja)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("LOJA");

                entity.Property(e => e.Natureza)
                    .HasMaxLength(15)
                    .IsUnicode(false)
                    .HasColumnName("NATUREZA");

                entity.Property(e => e.Vlorcado).HasColumnName("VLORCADO");
            });

            modelBuilder.Entity<PermClisec>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("PERM_CLISEC");

                entity.Property(e => e.Codcliprin)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CODCLIPRIN");

                entity.Property(e => e.Codclisec)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CODCLISEC");

                entity.Property(e => e.Lojacliprin)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("LOJACLIPRIN");

                entity.Property(e => e.Lojaclisec)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("LOJACLISEC");
            });

            modelBuilder.Entity<Resumorlt>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("RESUMORLT");

                entity.Property(e => e.Anomes)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("ANOMES");

                entity.Property(e => e.Bdgfat).HasColumnName("BDGFAT");

                entity.Property(e => e.Bdgqtdc).HasColumnName("BDGQTDC");

                entity.Property(e => e.Bdgval).HasColumnName("BDGVAL");

                entity.Property(e => e.Empresa)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("EMPRESA");

                entity.Property(e => e.Faturado).HasColumnName("FATURADO");

                entity.Property(e => e.Qtdcir).HasColumnName("QTDCIR");

                entity.Property(e => e.Tipocli)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("TIPOCLI");

                entity.Property(e => e.Valorizado).HasColumnName("VALORIZADO");
            });

            modelBuilder.Entity<Resumoval>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("RESUMOVAL");

                entity.Property(e => e.Anomes)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("ANOMES");

                entity.Property(e => e.Empresa)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("EMPRESA");

                entity.Property(e => e.Faturado).HasColumnName("FATURADO");

                entity.Property(e => e.Orcado).HasColumnName("ORCADO");

                entity.Property(e => e.Tipocli)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("TIPOCLI");

                entity.Property(e => e.Valorizado).HasColumnName("VALORIZADO");
            });

            modelBuilder.Entity<TbLicitIten>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("TB_LICIT_ITENS");

                entity.Property(e => e.Codcli)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("CODCLI");

                entity.Property(e => e.Descitem)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("DESCITEM");

                entity.Property(e => e.Empenho)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("EMPENHO");

                entity.Property(e => e.Grupo).HasColumnName("GRUPO");

                entity.Property(e => e.Item)
                    .HasMaxLength(10)
                    .IsUnicode(false)
                    .HasColumnName("ITEM");

                entity.Property(e => e.Lojacli)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("LOJACLI");

                entity.Property(e => e.Numproc)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("NUMPROC");

                entity.Property(e => e.Qtde).HasColumnName("QTDE");

                entity.Property(e => e.Vlunit).HasColumnName("VLUNIT");
            });

            modelBuilder.Entity<TbLicitProd>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("TB_LICIT_PROD");

                entity.Property(e => e.Codprod)
                    .HasMaxLength(15)
                    .IsUnicode(false)
                    .HasColumnName("CODPROD");

                entity.Property(e => e.Empenho)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("EMPENHO");

                entity.Property(e => e.Grupo).HasColumnName("GRUPO");

                entity.Property(e => e.Item)
                    .HasMaxLength(10)
                    .IsUnicode(false)
                    .HasColumnName("ITEM");

                entity.Property(e => e.Numproc)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("NUMPROC");

                entity.Property(e => e.Princ)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("PRINC")
                    .IsFixedLength();
            });

            modelBuilder.Entity<Tempd2rlt>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("TEMPD2RLT");

                entity.Property(e => e.Emissao)
                    .HasMaxLength(8)
                    .IsUnicode(false)
                    .HasColumnName("EMISSAO");

                entity.Property(e => e.Filial)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("FILIAL");

                entity.Property(e => e.Nf)
                    .HasMaxLength(9)
                    .IsUnicode(false)
                    .HasColumnName("NF");

                entity.Property(e => e.Pedido)
                    .HasMaxLength(6)
                    .IsUnicode(false)
                    .HasColumnName("PEDIDO");

                entity.Property(e => e.Serie)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("SERIE");

                entity.Property(e => e.Total).HasColumnName("TOTAL");
            });

            modelBuilder.Entity<Validint>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("VALIDINT");

                entity.Property(e => e.Codprod)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("CODPROD");

                entity.Property(e => e.Descprod)
                    .HasMaxLength(100)
                    .IsUnicode(false)
                    .HasColumnName("DESCPROD");

                entity.Property(e => e.Lote)
                    .HasMaxLength(30)
                    .IsUnicode(false)
                    .HasColumnName("LOTE");
            });

            modelBuilder.Entity<Validint2>(entity =>
            {
                entity.HasNoKey();

                entity.ToTable("VALIDINT2");

                entity.Property(e => e.Codprod)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("CODPROD");

                entity.Property(e => e.Descprod)
                    .HasMaxLength(100)
                    .IsUnicode(false)
                    .HasColumnName("DESCPROD");

                entity.Property(e => e.Lote)
                    .HasMaxLength(30)
                    .IsUnicode(false)
                    .HasColumnName("LOTE");
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
