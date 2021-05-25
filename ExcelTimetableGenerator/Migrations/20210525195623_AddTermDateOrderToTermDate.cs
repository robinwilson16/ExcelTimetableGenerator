using Microsoft.EntityFrameworkCore.Migrations;

namespace ExcelTimetableGenerator.Migrations
{
    public partial class AddTermDateOrderToTermDate : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "TermDateOrder",
                table: "ETG_TermDate",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "TermDateOrder",
                table: "ETG_TermDate");
        }
    }
}
