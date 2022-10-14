using Microsoft.EntityFrameworkCore;
using System;
using TestCaseAnalyzer.App.Database;

namespace TestCaseAnalyzer.App
{
    class Program
    {
        private const string DbConnectionString = "Server=.\\MYSQLSERVER;Database=TestCaseAnalyzer;Integrated Security=True;Persist Security Info=False;Pooling=False;MultipleActiveResultSets=False;Connect Timeout=60;Encrypt=False;TrustServerCertificate=False";

        static void Main(string[] args)
        {
            var options = new DbContextOptionsBuilder()
                .UseSqlServer(DbConnectionString)
                .Options;

            using var db = new TestCaseAnalyzerDbContext(options);

            db.

            db.TestCaseItems.Add(new Database.Entities.TestCaseItem
            {
                Objective = "helo",
                TestCaseId = "gobye"
            });

            db.SaveChanges();


            return;
            App.RunMyApp();
        }
    
    }
}
