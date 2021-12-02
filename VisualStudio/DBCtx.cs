using MyExel.Models;
using Microsoft.EntityFrameworkCore;

namespace MyExel
{
    public class DBCtx : DbContext
    {
        public DBCtx(DbContextOptions<DBCtx> options) : base(options)
        {
        }
        public DbSet<med> med { get; set; }
        public DbSet<mkb> mkb { get; set; }
        public DbSet<mo> mo { get; set; }
        public DbSet<smo> smo { get; set; }
    }
}