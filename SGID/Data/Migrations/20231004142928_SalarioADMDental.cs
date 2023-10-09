using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class SalarioADMDental : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<double>(
                name: "Salario",
                table: "TimeDentals",
                type: "float",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<double>(
                name: "Teto",
                table: "TimeDentals",
                type: "float",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<double>(
                name: "Salario",
                table: "TimeADMs",
                type: "float",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<double>(
                name: "Teto",
                table: "TimeADMs",
                type: "float",
                nullable: false,
                defaultValue: 0.0);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Salario",
                table: "TimeDentals");

            migrationBuilder.DropColumn(
                name: "Teto",
                table: "TimeDentals");

            migrationBuilder.DropColumn(
                name: "Salario",
                table: "TimeADMs");

            migrationBuilder.DropColumn(
                name: "Teto",
                table: "TimeADMs");
        }
    }
}
