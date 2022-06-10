USE [PPTL]
GO

/****** Object:  Table [dbo].[Z_Sales_Written_Manager_Principal_Product_Qty]    Script Date: 10/02/2021 11:13:01 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Z_Sales_Written_Manager_Principal_Product_Qty](
	[Manager] [varchar](max) NULL,
	[Principal] [varchar](max) NULL,
	[Product] [varchar](max) NULL,
	[MontylyTargetQty] [bigint] NULL,
	[PricePerUnit] [decimal](18, 2) NULL,
	[MonthlyTargetVal] [decimal](18, 2) NULL,
	[2019_Sales] [bigint] NULL,
	[2020_Sales] [bigint] NULL,
	[Jan] [bigint] NULL,
	[Feb] [bigint] NULL,
	[Mar] [bigint] NULL,
	[Apr] [bigint] NULL,
	[May] [bigint] NULL,
	[Jun] [bigint] NULL,
	[Jul] [bigint] NULL,
	[Aug] [bigint] NULL,
	[Sep] [bigint] NULL,
	[Oct] [bigint] NULL,
	[Nov] [bigint] NULL,
	[Dec] [bigint] NULL,
	[GrandTotal] [bigint] NULL,
	[TargetTillLastMonth] [bigint] NULL,
	[SalesTillLastMonth] [bigint] NULL,
	[Revised_Target] [bigint] NULL,
	[Jan2020Qty] [bigint] NULL,
	[Feb2020Qty] [bigint] NULL,
	[Mar2020Qty] [bigint] NULL,
	[Apr2020Qty] [bigint] NULL,
	[May2020Qty] [bigint] NULL,
	[Jun2020Qty] [bigint] NULL,
	[Jul2020Qty] [bigint] NULL,
	[Aug2020Qty] [bigint] NULL,
	[Sep2020Qty] [bigint] NULL,
	[Oct2020Qty] [bigint] NULL,
	[Nov2020Qty] [bigint] NULL,
	[Dec2020Qty] [bigint] NULL,
	[GrandTotalQty] [bigint] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


