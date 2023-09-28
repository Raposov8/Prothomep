using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class DentalComissões : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "TimeDentals",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    DataCriacao = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Status = table.Column<bool>(type: "bit", nullable: false),
                    Integrante = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Lider = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Meta = table.Column<double>(type: "float", nullable: false),
                    Porcentagem = table.Column<double>(type: "float", nullable: false),
                    PorcentagemSeg = table.Column<double>(type: "float", nullable: false),
                    IdUsuario = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Color = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_TimeDentals", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "TimeDentals");
        }
    }
}
