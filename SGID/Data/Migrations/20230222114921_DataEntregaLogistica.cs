using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class DataEntregaLogistica : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<bool>(
                name: "GerenProd",
                table: "Times",
                type: "bit",
                nullable: false,
                defaultValue: false);

            migrationBuilder.AddColumn<double>(
                name: "PorcentagemGenProd",
                table: "Times",
                type: "float",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<DateTime>(
                name: "DataEntrega",
                table: "Agendamentos",
                type: "datetime2",
                nullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "GerenProd",
                table: "Times");

            migrationBuilder.DropColumn(
                name: "PorcentagemGenProd",
                table: "Times");

            migrationBuilder.DropColumn(
                name: "DataEntrega",
                table: "Agendamentos");
        }
    }
}
