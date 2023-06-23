using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class Usuario : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "UsuarioAlterar",
                table: "Cotacoes",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioCriacao",
                table: "Cotacoes",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioAlterar",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioCriacao",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "UsuarioAlterar",
                table: "Cotacoes");

            migrationBuilder.DropColumn(
                name: "UsuarioCriacao",
                table: "Cotacoes");

            migrationBuilder.DropColumn(
                name: "UsuarioAlterar",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "UsuarioCriacao",
                table: "Agendamentos");
        }
    }
}
