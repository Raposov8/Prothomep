using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class TimePJ : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<bool>(
                name: "PJ",
                table: "Times",
                type: "bit",
                nullable: false,
                defaultValue: false);

            migrationBuilder.AddColumn<bool>(
                name: "PJ",
                table: "TimeDentals",
                type: "bit",
                nullable: false,
                defaultValue: false);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "PJ",
                table: "Times");

            migrationBuilder.DropColumn(
                name: "PJ",
                table: "TimeDentals");
        }
    }
}
