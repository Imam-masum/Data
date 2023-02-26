using Microsoft.EntityFrameworkCore;

namespace DataExcel.Models
{
    public class AppDbContext:DbContext
    {
        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }
        public DbSet<Country> Countries { get; set; }
    }
}
