USE [PPTL]
GO

/****** Object:  Table [dbo].[Z_Sales_Written_Manager_Principal_Particular_Product_Qty]    Script Date: 10/02/2021 11:12:51 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Z_Sales_Written_Manager_Principal_Particular_Product_Qty](
	[Manager] [varchar](800) NULL,
	[Particulars] [varchar](800) NULL,
	[Principal] [varchar](800) NULL,
	[Product] [varchar](800) NULL,
	[GrandTotalQty] [bigint] NULL,
	[JanQty] [bigint] NULL,
	[FebQty] [bigint] NULL,
	[MarQty] [bigint] NULL,
	[AprQty] [bigint] NULL,
	[MayQty] [bigint] NULL,
	[JunQty] [bigint] NULL,
	[JulQty] [bigint] NULL,
	[AugQty] [bigint] NULL,
	[SepQty] [bigint] NULL,
	[OctQty] [bigint] NULL,
	[NovQty] [bigint] NULL,
	[DecQty] [bigint] NULL,
	[GrandTotalQty2020] [bigint] NULL,
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
	[GrandTotalQty2021] [bigint] NULL,
	[Jan21Target] [bigint] NULL,
	[Feb21Target] [bigint] NULL,
	[Mar21Target] [bigint] NULL,
	[Apr21Target] [bigint] NULL,
	[May21Target] [bigint] NULL,
	[Jun21Target] [bigint] NULL,
	[Jul21Target] [bigint] NULL,
	[Aug21Target] [bigint] NULL,
	[Sep21Target] [bigint] NULL,
	[Oct21Target] [bigint] NULL,
	[Nov21Target] [bigint] NULL,
	[Dec21Target] [bigint] NULL
) ON [PRIMARY]
GO


