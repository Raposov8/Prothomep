using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class DadosCirurgia : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "DadosCirurgias",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DataCriacao = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Codigo = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    NomePaciente = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    NomeMedico = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    NomeCliente = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Status = table.Column<int>(type: "int", nullable: false),
                    DataCirurgia = table.Column<DateTime>(type: "datetime2", nullable: false),
                    HoraCirurgia = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    ProcedimentosExec = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    ObsIntercorrencia = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_DadosCirurgias", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "AnexosDadosCirurgias",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DadosCirurgiaId = table.Column<int>(type: "int", nullable: false),
                    AnexoCam = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_AnexosDadosCirurgias", x => x.Id);
                    table.ForeignKey(
                        name: "FK_AnexosDadosCirurgias_DadosCirurgias_DadosCirurgiaId",
                        column: x => x.DadosCirurgiaId,
                        principalTable: "DadosCirurgias",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_AnexosDadosCirurgias_DadosCirurgiaId",
                table: "AnexosDadosCirurgias",
                column: "DadosCirurgiaId");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "AnexosDadosCirurgias");

            migrationBuilder.DropTable(
                name: "DadosCirurgias");
        }
    }
}
