using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class RecursoFK : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "BionexoCategorias");

            migrationBuilder.DropTable(
                name: "BionexoUnidadeMedidas");

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
                name: "AnexosCotacoes",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    CotacaoId = table.Column<int>(type: "int", nullable: false),
                    AnexosCam = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_AnexosCotacoes", x => x.Id);
                    table.ForeignKey(
                        name: "FK_AnexosCotacoes_Cotacoes_CotacaoId",
                        column: x => x.CotacaoId,
                        principalTable: "Cotacoes",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_Agendamentos_InstrumentadorId",
                table: "Agendamentos",
                column: "InstrumentadorId");

            migrationBuilder.CreateIndex(
                name: "IX_AnexosCotacoes_CotacaoId",
                table: "AnexosCotacoes",
                column: "CotacaoId");

            migrationBuilder.AddForeignKey(
                name: "FK_Agendamentos_Recursos_InstrumentadorId",
                table: "Agendamentos",
                column: "InstrumentadorId",
                principalTable: "Recursos",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_Agendamentos_Recursos_InstrumentadorId",
                table: "Agendamentos");

            migrationBuilder.DropTable(
                name: "AnexosCotacoes");

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

            migrationBuilder.CreateTable(
                name: "BionexoCategorias",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Descricao = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BionexoCategorias", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "BionexoUnidadeMedidas",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Abreviacao = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Descricao = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BionexoUnidadeMedidas", x => x.Id);
                });
        }
    }
}
