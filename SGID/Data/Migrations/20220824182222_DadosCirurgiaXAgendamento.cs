using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class DadosCirurgiaXAgendamento : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "AgendamentoId",
                table: "DadosCirurgias",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.CreateIndex(
                name: "IX_DadosCirurgias_AgendamentoId",
                table: "DadosCirurgias",
                column: "AgendamentoId");

            migrationBuilder.AddForeignKey(
                name: "FK_DadosCirurgias_Agendamentos_AgendamentoId",
                table: "DadosCirurgias",
                column: "AgendamentoId",
                principalTable: "Agendamentos",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_DadosCirurgias_Agendamentos_AgendamentoId",
                table: "DadosCirurgias");

            migrationBuilder.DropIndex(
                name: "IX_DadosCirurgias_AgendamentoId",
                table: "DadosCirurgias");

            migrationBuilder.DropColumn(
                name: "AgendamentoId",
                table: "DadosCirurgias");
        }
    }
}
