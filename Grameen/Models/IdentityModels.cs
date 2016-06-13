using System.Data.Entity;
using Microsoft.AspNet.Identity.EntityFramework;
using Optimize;
using Database = System.Data.Entity.Database;

namespace Grameen.Models
{
    // You can add profile data for the user by adding more properties to your ApplicationUser class, please visit http://go.microsoft.com/fwlink/?LinkID=317594 to learn more.
    public class ApplicationUser : IdentityUser
    {
    }

    public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
    {
        public ApplicationDbContext()
            : base("DefaultConnection")
        {
            //disable initializer
            Database.SetInitializer<ApplicationDbContext>(null);
        }

        public DbSet<Crop> Crops { get; set; }
        public DbSet<Region> Regions { get; set; }
        public DbSet<RegionCrop> RegionCrops { get; set; }
        public DbSet<Activity> Activities { get; set; }
        public DbSet<Error> Errors { get; set; }
    }
}