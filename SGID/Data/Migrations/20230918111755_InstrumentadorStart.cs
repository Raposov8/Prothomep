using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class InstrumentadorStart : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "AgendamentoInstrumentadors",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Data = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Instrumentador = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Agendamento = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Procedimento = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Hospital = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Medico = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Tipo = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Status = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_AgendamentoInstrumentadors", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "AgendamentoInstrumentadors");
        }
    }
}
