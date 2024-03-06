using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class TipoComissao : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<double>(
                name: "AtingimentoMeta",
                table: "TimeDentals",
                type: "float",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<double>(
                name: "PorcentagemEtapaDois",
                table: "TimeDentals",
                type: "float",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<double>(
                name: "PorcentagemEtapaUm",
                table: "TimeDentals",
                type: "float",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<string>(
                name: "TipoComissao",
                table: "TimeDentals",
                type: "nvarchar(max)",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "AtingimentoMeta",
                table: "TimeDentals");

            migrationBuilder.DropColumn(
                name: "PorcentagemEtapaDois",
                table: "TimeDentals");

            migrationBuilder.DropColumn(
                name: "PorcentagemEtapaUm",
                table: "TimeDentals");

            migrationBuilder.DropColumn(
                name: "TipoComissao",
                table: "TimeDentals");
        }
    }
}
