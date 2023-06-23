using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class SGID06 : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_Cotacoes_AgendamentoId",
                table: "Cotacoes");

            migrationBuilder.AlterColumn<int>(
                name: "Tipo",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                oldClrType: typeof(string),
                oldType: "nvarchar(max)");

            migrationBuilder.AddColumn<int>(
                name: "Autorizado",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "CotacaoId",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<DateTime>(
                name: "DataAlteracao",
                table: "Agendamentos",
                type: "datetime2",
                nullable: false,
                defaultValue: new DateTime(1, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified));

            migrationBuilder.AddColumn<string>(
                name: "Empresa",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.CreateIndex(
                name: "IX_Cotacoes_AgendamentoId",
                table: "Cotacoes",
                column: "AgendamentoId",
                unique: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_Cotacoes_AgendamentoId",
                table: "Cotacoes");

            migrationBuilder.DropColumn(
                name: "Autorizado",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "CotacaoId",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "DataAlteracao",
                table: "Agendamentos");

            migrationBuilder.DropColumn(
                name: "Empresa",
                table: "Agendamentos");

            migrationBuilder.AlterColumn<string>(
                name: "Tipo",
                table: "Agendamentos",
                type: "nvarchar(max)",
                nullable: false,
                oldClrType: typeof(int),
                oldType: "int");

            migrationBuilder.CreateIndex(
                name: "IX_Cotacoes_AgendamentoId",
                table: "Cotacoes",
                column: "AgendamentoId");
        }
    }
}
