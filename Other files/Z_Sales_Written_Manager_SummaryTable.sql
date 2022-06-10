USE [PPTL]
GO

/****** Object:  Table [dbo].[Z_Sales_Written_Manager_SummaryTable]    Script Date: 10/02/2021 11:13:05 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Z_Sales_Written_Manager_SummaryTable](
	[Manager] [varchar](max) NULL,
	[Principal] [varchar](max) NULL,
	[Target] [bigint] NULL,
	[2019Sales] [bigint] NULL,
	[2020Sales] [bigint] NULL,
	[Jan_Val] [decimal](18, 2) NULL,
	[Feb_Val] [decimal](18, 2) NULL,
	[Mar_Val] [decimal](18, 2) NULL,
	[Apr_Val] [decimal](18, 2) NULL,
	[May_Val] [decimal](18, 2) NULL,
	[Jun_Val] [decimal](18, 2) NULL,
	[Jul_Val] [decimal](18, 2) NULL,
	[Aug_Val] [decimal](18, 2) NULL,
	[Sep_Val] [decimal](18, 2) NULL,
	[Oct_Val] [decimal](18, 2) NULL,
	[Nov_Val] [decimal](18, 2) NULL,
	[Dec_Val] [decimal](18, 2) NULL,
	[TargettilllastMonth] [bigint] NULL,
	[SalestilllastMonth] [decimal](18, 2) NULL,
	[Achievement] [decimal](18, 2) NULL,
	[GrandTotalVal] [decimal](18, 2) NULL,
	[Revised_Target] [decimal](18, 2) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


