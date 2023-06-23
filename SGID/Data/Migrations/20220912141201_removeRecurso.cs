using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class removeRecurso : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "DataCirurgiaFinal",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Hora",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "HoraFinal",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Loja",
                table: "Agendamentos");

            migrationBuilder.AddColumn<double>(
                name: "ValorTotal",
                table: "ProdutosAgendamentos",
                type: "float",
                nullable: false,
                defaultValue: 0.0);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "ValorTotal",
                table: "ProdutosAgendamentos");

            migrationBuilder.AddColumn<DateTime>(
                name: "DataCirurgiaFinal",
                table: "Agendamentos",
                type: "datetime2",
                nullable: false,
                defaultValue: new DateTime(1, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified));

            migrationBuilder.AddColumn<string>(
                name: "Hora",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "HoraFinal",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Loja",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");
        }
    }
}
