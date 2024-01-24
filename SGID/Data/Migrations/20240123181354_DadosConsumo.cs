using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class DadosConsumo : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "DadosCirugiasProdutos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DadosCirurgiaId = table.Column<int>(type: "int", nullable: false),
                    Produto = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Quantidade = table.Column<double>(type: "float", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_DadosCirugiasProdutos", x => x.Id);
                    table.ForeignKey(
                        name: "FK_DadosCirugiasProdutos_DadosCirurgias_DadosCirurgiaId",
                        column: x => x.DadosCirurgiaId,
                        principalTable: "DadosCirurgias",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_DadosCirugiasProdutos_DadosCirurgiaId",
                table: "DadosCirugiasProdutos",
                column: "DadosCirurgiaId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "DadosCirugiasProdutos");
        }
    }
}
