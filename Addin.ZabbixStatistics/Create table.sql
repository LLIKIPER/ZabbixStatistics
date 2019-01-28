CREATE TABLE [dbo].[Statistics](
	[Time] [datetime] NOT NULL,
	[Basename] [nchar](50) NOT NULL,
	[Param] [nchar](50) NOT NULL,
	[Value] [int] NOT NULL
) ON [PRIMARY]