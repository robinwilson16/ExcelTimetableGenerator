using Microsoft.EntityFrameworkCore.Migrations;

namespace ExcelTimetableGenerator.Migrations
{
    public partial class FirstMigration : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "ETG_BankHoliday",
                columns: table => new
                {
                    BankHolidayID = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    AcademicYearID = table.Column<string>(type: "nvarchar(5)", maxLength: 5, nullable: true),
                    BankHolidayDesc = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ETG_BankHoliday", x => x.BankHolidayID);
                });

            migrationBuilder.CreateTable(
                name: "ETG_DaySlot",
                columns: table => new
                {
                    DaySlotID = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    AcademicYearID = table.Column<string>(type: "nvarchar(5)", maxLength: 5, nullable: true),
                    DayName = table.Column<string>(type: "nvarchar(20)", maxLength: 20, nullable: true),
                    Slots = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ETG_DaySlot", x => x.DaySlotID);
                });

            migrationBuilder.CreateTable(
                name: "ETG_TermDate",
                columns: table => new
                {
                    TermDateID = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    AcademicYearID = table.Column<string>(type: "nvarchar(5)", maxLength: 5, nullable: true),
                    TermDateName = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    IsTerm = table.Column<bool>(type: "bit", nullable: false),
                    Dates = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ETG_TermDate", x => x.TermDateID);
                });

            migrationBuilder.CreateTable(
                name: "ETG_TimetableSection",
                columns: table => new
                {
                    TimetableSectionID = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    SectionName = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ETG_TimetableSection", x => x.TimetableSectionID);
                });

            migrationBuilder.CreateTable(
                name: "ETG_Week",
                columns: table => new
                {
                    WeekID = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    AcademicYearID = table.Column<string>(type: "nvarchar(5)", maxLength: 5, nullable: true),
                    WeekNum = table.Column<int>(type: "int", nullable: true),
                    WeekDesc = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    Notes = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: true),
                    Notes2 = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ETG_Week", x => x.WeekID);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "ETG_BankHoliday");

            migrationBuilder.DropTable(
                name: "ETG_DaySlot");

            migrationBuilder.DropTable(
                name: "ETG_TermDate");

            migrationBuilder.DropTable(
                name: "ETG_TimetableSection");

            migrationBuilder.DropTable(
                name: "ETG_Week");
        }
    }
}
