using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class FormularioAvulso : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "FormularioAvulsos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DataCriacao = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Cliente = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    DataCirurgia = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Paciente = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Cirurgiao = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Convenio = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Representante = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    NumAgendamento = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Usuario = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Empresa = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_FormularioAvulsos", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "FormularioAvulsoXProdutos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    FormularioId = table.Column<int>(type: "int", nullable: false),
                    Produto = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Descricao = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Quantidade = table.Column<double>(type: "float", nullable: false),
                    Lote = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_FormularioAvulsoXProdutos", x => x.Id);
                    table.ForeignKey(
                        name: "FK_FormularioAvulsoXProdutos_FormularioAvulsos_FormularioId",
                        column: x => x.FormularioId,
                        principalTable: "FormularioAvulsos",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_FormularioAvulsoXProdutos_FormularioId",
                table: "FormularioAvulsoXProdutos",
                column: "FormularioId");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "FormularioAvulsoXProdutos");

            migrationBuilder.DropTable(
                name: "FormularioAvulsos");
        }
    }
}
