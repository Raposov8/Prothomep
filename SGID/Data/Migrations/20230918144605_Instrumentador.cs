using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SGID.Data.Migrations
{
    /// <inheritdoc />
    public partial class Instrumentador : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Instrumentadores",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Nome = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Nascimento = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Endereco = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Bairro = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Municipio = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Estado = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Pais = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Email = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Telefone = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Telefone2 = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    CBO = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    RG = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    CPF = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    PISPASEP = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    TipoChave = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Banco = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Ag = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    CC = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Remuneracao = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    SErvCon = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Instrumentadores", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Instrumentadores");
        }
    }
}
