using System;
using Microsoft.EntityFrameworkCore.Migrations;

namespace GroupsConnector.Migrations
{
    public partial class InitialCreate : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Groups",
                columns: table => new
                {
                    Id = table.Column<string>(nullable: false),
                    DisplayName = table.Column<string>(nullable: true),
                    Description = table.Column<string>(nullable: true),
                    IsDeleted = table.Column<bool>(nullable: false, defaultValue: false),
                    LastUpdated = table.Column<DateTime>(nullable: false, defaultValueSql: "datetime()")
                });

            // Set trigger to update the LastUpdated property
            migrationBuilder.Sql(@"CREATE TRIGGER set_last_updated AFTER UPDATE
            ON Groups
            BEGIN
                UPDATE Groups
                SET LastUpdated = datetime('now')
                WHERE Id = NEW.Id;
            END;");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql("DROP TRIGGER set_last_updated");

            migrationBuilder.DropTable(
                name: "Groups");
        }
    }
}
