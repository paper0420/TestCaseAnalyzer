CREATE TABLE [dbo].[TestCaseItem]
(
	[Id] INT NOT NULL IDENTITY(1, 1) PRIMARY KEY, 
    [TestCaseId] VARCHAR(20) NOT NULL, 
    [Objective] VARCHAR(MAX) NOT NULL 
)
