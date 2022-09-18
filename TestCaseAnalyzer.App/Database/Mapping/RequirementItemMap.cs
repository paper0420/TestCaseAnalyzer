using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using TestCaseAnalyzer.App.Database.Entities;

namespace TestCaseAnalyzer.App.Database.Mapping
{
    internal class RequirementItemMap : IEntityTypeConfiguration<RequirementItem>
    {
        public void Configure(EntityTypeBuilder<RequirementItem> builder)
        {
            builder.ToTable("RequirementItem");
            builder.HasKey(t => t.Id);
            builder.Property(t => t.Id).HasColumnName("Id");
            builder.Property(t => t.TestCaseId).HasColumnName("TestCaseId");
            builder.Property(t => t.Objective).HasColumnName("Objective");
        }
    }
}
