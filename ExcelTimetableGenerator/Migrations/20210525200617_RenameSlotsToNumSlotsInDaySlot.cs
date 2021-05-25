using Microsoft.EntityFrameworkCore.Migrations;

namespace ExcelTimetableGenerator.Migrations
{
    public partial class RenameSlotsToNumSlotsInDaySlot : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "Slots",
                table: "ETG_DaySlot",
                newName: "NumSlots");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "NumSlots",
                table: "ETG_DaySlot",
                newName: "Slots");
        }
    }
}
