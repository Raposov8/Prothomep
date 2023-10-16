using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class DataLogsAgendamento : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "DataRetirada",
                table: "Agendamentos",
                newName: "DataRetorno");

            migrationBuilder.AddColumn<DateTime>(
                name: "DataComercial",
                table: "Agendamentos",
                type: "datetime2",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "DataComercialAprova",
                table: "Agendamentos",
                type: "datetime2",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "DataLogistica",
                table: "Agendamentos",
                type: "datetime2",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioComercialAprova",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "DataComercial",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "DataComercialAprova",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "DataLogistica",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "UsuarioComercialAprova",
                table: "Agendamentos");

            migrationBuilder.RenameColumn(
                name: "DataRetorno",
                table: "Agendamentos",
                newName: "DataRetirada");
        }
    }
}
