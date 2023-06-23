using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class remocaoRecursos : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_Agendamentos_Recursos_InstrumentadorId",
                table: "Agendamentos");

            migrationBuilder.DropTable(
                name: "Recursos");

            migrationBuilder.DropIndex(
                name: "IX_Agendamentos_InstrumentadorId",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "InstrumentadorId",
                table: "Agendamentos");

            migrationBuilder.AddColumn<string>(
                name: "Instrumentador",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Instrumentador",
                table: "Agendamentos");

            migrationBuilder.AddColumn<int>(
                name: "InstrumentadorId",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.CreateTable(
                name: "Recursos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    CPF = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Codigo = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    DELET = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    DataCriacao = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Imagem = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Nome = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Observação = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    RG = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Status = table.Column<int>(type: "int", nullable: false),
                    Tipo = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    TipoContrato = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Recursos", x => x.Id);
                });

            migrationBuilder.CreateIndex(
                name: "IX_Agendamentos_InstrumentadorId",
                table: "Agendamentos",
                column: "InstrumentadorId");

            migrationBuilder.AddForeignKey(
                name: "FK_Agendamentos_Recursos_InstrumentadorId",
                table: "Agendamentos",
                column: "InstrumentadorId",
                principalTable: "Recursos",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }
    }
}
