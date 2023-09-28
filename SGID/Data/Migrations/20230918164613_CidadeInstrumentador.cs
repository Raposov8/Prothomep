using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class CidadeInstrumentador : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "SErvCon",
                table: "Instrumentadores",
                newName: "ServCon");

            migrationBuilder.AddColumn<string>(
                name: "ChavePix",
                table: "Instrumentadores",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Empresa",
                table: "AgendamentoInstrumentadors",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "ChavePix",
                table: "Instrumentadores");

            migrationBuilder.DropColumn(
                name: "Empresa",
                table: "AgendamentoInstrumentadors");

            migrationBuilder.RenameColumn(
                name: "ServCon",
                table: "Instrumentadores",
                newName: "SErvCon");
        }
    }
}
