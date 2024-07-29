using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class ADMaprova : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<DateTime>(
                name: "DataAprova",
                table: "Agendamentos",
                type: "datetime2",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "UsuarioAprova",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "DataAprova",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "UsuarioAprova",
                table: "Agendamentos");
        }
    }
}
