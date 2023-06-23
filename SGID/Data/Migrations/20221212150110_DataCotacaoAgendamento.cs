using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    public partial class DataCotacaoAgendamento : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "ProcedimentoCotacaos");

            migrationBuilder.DropTable(
                name: "ProdutosCotacoes");

            migrationBuilder.DropTable(
                name: "Cotacoes");

            migrationBuilder.DropColumn(
                name: "CotacaoId",
                table: "Agendamentos");

            migrationBuilder.AddColumn<DateTime>(
                name: "DataCotacao",
                table: "Agendamentos",
                type: "datetime2",
                nullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "DataCotacao",
                table: "Agendamentos");

            migrationBuilder.AddColumn<int>(
                name: "CotacaoId",
                table: "Agendamentos",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.CreateTable(
                name: "Cotacoes",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    AgendamentoId = table.Column<int>(type: "int", nullable: false),
                    CRM = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    DataCirurgia = table.Column<DateTime>(type: "datetime2", nullable: false),
                    DataCriacao = table.Column<DateTime>(type: "datetime2", nullable: false),
                    NomeCliente = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    NomeHospital = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    NomeMedico = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    NomePaciente = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    NomeVendedor = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Status = table.Column<int>(type: "int", nullable: false),
                    Urgencia = table.Column<int>(type: "int", nullable: false),
                    UsuarioAlterar = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    UsuarioCriacao = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Cotacoes", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Cotacoes_Agendamentos_AgendamentoId",
                        column: x => x.AgendamentoId,
                        principalTable: "Agendamentos",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "ProcedimentoCotacaos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    CotacaoId = table.Column<int>(type: "int", nullable: false),
                    ProcedimentoId = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ProcedimentoCotacaos", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "ProdutosCotacoes",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    CotacaoId = table.Column<int>(type: "int", nullable: false),
                    Codigo = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Descricao = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Quantidade = table.Column<double>(type: "float", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ProdutosCotacoes", x => x.Id);
                    table.ForeignKey(
                        name: "FK_ProdutosCotacoes_Cotacoes_CotacaoId",
                        column: x => x.CotacaoId,
                        principalTable: "Cotacoes",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_Cotacoes_AgendamentoId",
                table: "Cotacoes",
                column: "AgendamentoId",
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_ProdutosCotacoes_CotacaoId",
                table: "ProdutosCotacoes",
                column: "CotacaoId");
        }
    }
}
