using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestCaseAnalyzer.App.Database.Entities;
using TestCaseAnalyzer.App.Database.Mapping;

namespace TestCaseAnalyzer.App.Database
{
    internal class TestCaseAnalyzerDbContext : DbContext
    {
        public TestCaseAnalyzerDbContext(DbContextOptions options) : base(options)
        {
        }

        public DbSet<TestCaseItem> TestCaseItems { get; set; }
        public DbSet<RequirementItem> RequirementItems { get; set; }


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.ApplyConfiguration(new TestCaseItemMap());
            modelBuilder.ApplyConfiguration(new RequirementItemMap());
        }
    }

}
