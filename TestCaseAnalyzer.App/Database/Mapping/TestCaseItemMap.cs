using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestCaseAnalyzer.App.Database.Entities;

namespace TestCaseAnalyzer.App.Database.Mapping
{
    internal class TestCaseItemMap : IEntityTypeConfiguration<TestCaseItem>
    {
        public void Configure(EntityTypeBuilder<TestCaseItem> builder)
        {
            builder.ToTable("TestCaseItem");
            builder.HasKey(t=>t.Id);
            builder.Property(t=>t.Id).HasColumnName("Id");
            builder.Property(t => t.TestCaseId).HasColumnName("TestCaseId");
            builder.Property(t => t.Objective).HasColumnName("Objective");
        }
    }
}
