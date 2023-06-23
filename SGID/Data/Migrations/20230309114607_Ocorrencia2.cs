using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class Ocorrencia2 : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Ocorrencias",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DataCriacao = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Empresa = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Cliente = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Medico = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Paciente = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Agendamento = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Cirurgia = table.Column<DateTime>(type: "datetime2", nullable: true),
                    Patrimonio = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    DescPatri = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Produto = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Descricao = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Quantidade = table.Column<int>(type: "int", nullable: false),
                    DataOcorrencia = table.Column<DateTime>(type: "datetime2", nullable: true),
                    Problema = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Acao = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Procedente = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Cobrado = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Reposto = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Vendedor = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Motorista = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Obs = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Ocorrencias", x => x.Id);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Ocorrencias");
        }
    }
}
