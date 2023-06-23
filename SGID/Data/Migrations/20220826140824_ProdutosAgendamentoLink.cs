using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class ProdutosAgendamentoLink : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateIndex(
                name: "IX_ProdutosAgendamentos_AgendamentoId",
                table: "ProdutosAgendamentos",
                column: "AgendamentoId");

            migrationBuilder.AddForeignKey(
                name: "FK_ProdutosAgendamentos_Agendamentos_AgendamentoId",
                table: "ProdutosAgendamentos",
                column: "AgendamentoId",
                principalTable: "Agendamentos",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_ProdutosAgendamentos_Agendamentos_AgendamentoId",
                table: "ProdutosAgendamentos");

            migrationBuilder.DropIndex(
                name: "IX_ProdutosAgendamentos_AgendamentoId",
                table: "ProdutosAgendamentos");
        }
    }
}
