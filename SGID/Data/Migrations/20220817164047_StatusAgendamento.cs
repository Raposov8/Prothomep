using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class StatusAgendamento : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "Liminar",
                table: "Cotacoes",
                newName: "Urgencia");

            migrationBuilder.RenameColumn(
                name: "CPFPaciente",
                table: "Cotacoes",
                newName: "NomeVendedor");

            migrationBuilder.AddColumn<string>(
                name: "NomeCliente",
                table: "Cotacoes",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "NomeHospital",
                table: "Cotacoes",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<int>(
                name: "StatusCotacao",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "StatusPedido",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "NomeCliente",
                table: "Cotacoes");

            migrationBuilder.DropColumn(
                name: "NomeHospital",
                table: "Cotacoes");

            migrationBuilder.DropColumn(
                name: "StatusCotacao",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "StatusPedido",
                table: "Agendamentos");

            migrationBuilder.RenameColumn(
                name: "Urgencia",
                table: "Cotacoes",
                newName: "Liminar");

            migrationBuilder.RenameColumn(
                name: "NomeVendedor",
                table: "Cotacoes",
                newName: "CPFPaciente");
        }
    }
}
