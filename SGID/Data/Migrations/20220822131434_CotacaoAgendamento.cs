using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class CotacaoAgendamento : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "AgendamentoId",
                table: "Cotacoes",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.CreateIndex(
                name: "IX_Cotacoes_AgendamentoId",
                table: "Cotacoes",
                column: "AgendamentoId");

            migrationBuilder.AddForeignKey(
                name: "FK_Cotacoes_Agendamentos_AgendamentoId",
                table: "Cotacoes",
                column: "AgendamentoId",
                principalTable: "Agendamentos",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_Cotacoes_Agendamentos_AgendamentoId",
                table: "Cotacoes");

            migrationBuilder.DropIndex(
                name: "IX_Cotacoes_AgendamentoId",
                table: "Cotacoes");

            migrationBuilder.DropColumn(
                name: "AgendamentoId",
                table: "Cotacoes");
        }
    }
}
