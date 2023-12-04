using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class InstrumentadorAgendamento : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<DateTime>(
                name: "DataInstrumentador",
                table: "Agendamentos",
                type: "datetime2",
                nullable: true);

            migrationBuilder.AddColumn<int>(
                name: "StatusInstrumentador",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioGestorInstrumentador",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioInstrumentador",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "DataInstrumentador",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "StatusInstrumentador",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "UsuarioGestorInstrumentador",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "UsuarioInstrumentador",
                table: "Agendamentos");
        }
    }
}
