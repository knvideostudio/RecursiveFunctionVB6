
if exists (select * from dbo.sysobjects where id = object_id(N'[tbCategoryMain]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [tbCategoryMain]
GO

CREATE TABLE [tbCategoryMain] (
	[CategoryID] [int] IDENTITY (1, 1) NOT NULL ,
	[UniqueValue] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[dtRegDate] [datetime] NULL CONSTRAINT [DF_tbCategoryMain_dtRegDate] DEFAULT (getdate()),
	CONSTRAINT [IX_tbCategoryMain] UNIQUE  NONCLUSTERED 
	(
		[UniqueValue]
	) WITH  FILLFACTOR = 75  ON [PRIMARY] 
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[tbCategoryRelation]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [tbCategoryRelation]
GO

CREATE TABLE [tbCategoryRelation] (
	[CategoryRelId] [int] IDENTITY (1, 1) NOT NULL ,
	[ParentUniqueValue] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ChildUniqueValue] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[dtRegDate] [datetime] NULL CONSTRAINT [DF_tbCategoryRelation_dtRegDate] DEFAULT (getdate())
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[tbCategoryText]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [tbCategoryText]
GO

CREATE TABLE [tbCategoryText] (
	[CategoryTextID] [int] IDENTITY (1, 1) NOT NULL ,
	[CategoryMainUniqueValue] [varchar] (16) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[sText] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[dtRegDate] [datetime] NULL CONSTRAINT [DF_tbCategoryText_dtRegDate] DEFAULT (getdate())
) ON [PRIMARY]
GO

