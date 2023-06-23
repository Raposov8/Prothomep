using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class BionexoUnidades : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropPrimaryKey(
                name: "PK_BionexoUnidadeMedida",
                table: "BionexoUnidadeMedida");

            migrationBuilder.RenameTable(
                name: "BionexoUnidadeMedida",
                newName: "BionexoUnidadeMedidas");

            migrationBuilder.AddPrimaryKey(
                name: "PK_BionexoUnidadeMedidas",
                table: "BionexoUnidadeMedidas",
                column: "Id");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropPrimaryKey(
                name: "PK_BionexoUnidadeMedidas",
                table: "BionexoUnidadeMedidas");

            migrationBuilder.RenameTable(
                name: "BionexoUnidadeMedidas",
                newName: "BionexoUnidadeMedida");

            migrationBuilder.AddPrimaryKey(
                name: "PK_BionexoUnidadeMedida",
                table: "BionexoUnidadeMedida",
                column: "Id");
        }
    }
}
