using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class NomeVendedor : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "IdRecurso",
                table: "AspNetUsers");

            migrationBuilder.AddColumn<string>(
                name: "NomeVendedor",
                table: "AspNetUsers",
                type: "nvarchar(max)",
                nullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "NomeVendedor",
                table: "AspNetUsers");

            migrationBuilder.AddColumn<int>(
                name: "IdRecurso",
                table: "AspNetUsers",
                type: "int",
                nullable: true);
        }
    }
}
