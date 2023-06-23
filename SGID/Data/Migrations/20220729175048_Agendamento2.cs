using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class Agendamento2 : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "IdHospital",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "IdInstrumentador",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "IdPaciente",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Status",
                table: "Agendamentos");

            migrationBuilder.AddColumn<string>(
                name: "Cliente",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Codigo",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "CondPag",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Convenio",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<DateTime>(
                name: "DataAutorizacao",
                table: "Agendamentos",
                type: "datetime2",
                nullable: false,
                defaultValue: new DateTime(2022, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified));

            migrationBuilder.AddColumn<DateTime>(
                name: "DataCirurgiaFinal",
                table: "Agendamentos",
                type: "datetime2",
                nullable: false,
                defaultValue: new DateTime(2022, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified));

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
                name: "Instrumentador",
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

            migrationBuilder.AddColumn<string>(
                name: "Matricula",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Medico",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "NumAutorizacao",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Observacao",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Paciente",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Senha",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Tabela",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Tipo",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "TipoCliente",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "Vendedor",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Cliente",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Codigo",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "CondPag",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Convenio",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "DataAutorizacao",
                table: "Agendamentos");

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
                name: "Instrumentador",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Loja",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Matricula",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Medico",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "NumAutorizacao",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Observacao",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Paciente",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Senha",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Tabela",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Tipo",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "TipoCliente",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Vendedor",
                table: "Agendamentos");

            migrationBuilder.AddColumn<int>(
                name: "IdHospital",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "IdInstrumentador",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "IdPaciente",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "Status",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }
    }
}
