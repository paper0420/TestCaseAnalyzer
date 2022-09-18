CREATE TABLE [dbo].[TestCaseItem]
(
	[Id] INT NOT NULL PRIMARY KEY, 
    [TestCaseId] VARCHAR(20) NOT NULL, 
    [Objective] VARCHAR(MAX) NOT NULL 
)
