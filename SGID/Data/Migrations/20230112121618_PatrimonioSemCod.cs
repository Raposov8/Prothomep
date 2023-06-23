using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class PatrimonioSemCod : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "CodigoPatrimonio",
                table: "PatrimoniosAgendamentos",
                newName: "Patrimonio");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "Patrimonio",
                table: "PatrimoniosAgendamentos",
                newName: "CodigoPatrimonio");
        }
    }
}
