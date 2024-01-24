using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class DadosCirurgiaRemoveHora : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "HoraCirurgia",
                table: "DadosCirurgias");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "HoraCirurgia",
                table: "DadosCirurgias",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");
        }
    }
}
