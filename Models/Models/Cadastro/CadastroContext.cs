using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace SGID.Models.Cadastro
{
    public partial class CadastroContext : DbContext
    {
        public CadastroContext()
        {
        }

        public CadastroContext(DbContextOptions<CadastroContext> options)
            : base(options)
        {
        }

        public virtual DbSet<Telefone> Telefones { get; set; } = null!;
        public virtual DbSet<Usuario> Usuarios { get; set; } = null!;

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Telefone>(entity =>
            {
                entity.ToTable("Telefone");

                entity.Property(e => e.Celular)
                    .HasMaxLength(50)
                    .IsFixedLength();

                entity.Property(e => e.Idcelular)
                    .HasMaxLength(50)
                    .IsFixedLength();

                entity.HasOne(d => d.Usuario)
                    .WithMany(p => p.Telefones)
                    .HasForeignKey(d => d.UsuarioId)
                    .HasConstraintName("fk_UsuarioId");
            });

            modelBuilder.Entity<Usuario>(entity =>
            {
                entity.Property(e => e.Email)
                    .HasMaxLength(50)
                    .IsFixedLength();

                entity.Property(e => e.Nome)
                    .HasMaxLength(50)
                    .IsFixedLength();

                entity.Property(e => e.Senha)
                    .HasMaxLength(50)
                    .IsFixedLength();
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
