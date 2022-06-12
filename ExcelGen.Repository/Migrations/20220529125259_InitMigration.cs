using Microsoft.EntityFrameworkCore.Migrations;

namespace ExcelGen.Repository.Migrations
{
    public partial class InitMigration : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Discipline",
                columns: table => new
                {
                    Id = table.Column<string>(nullable: false),
                    Code = table.Column<string>(nullable: true),
                    Name = table.Column<string>(nullable: true),
                    ControlType = table.Column<int>(nullable: false),
                    NumberOfECTS = table.Column<int>(nullable: false),
                    LectureHours = table.Column<int>(nullable: false),
                    LabHours = table.Column<int>(nullable: false),
                    PracticeHours = table.Column<int>(nullable: false),
                    Category = table.Column<int>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Discipline", x => x.Id);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Discipline");
        }
    }
}
