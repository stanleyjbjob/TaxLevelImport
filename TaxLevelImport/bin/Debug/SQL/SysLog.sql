CREATE TABLE [dbo].[SYSLOG](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[TYPE] [nvarchar](50) NOT NULL,
	[SOURCE] [nvarchar](50) NOT NULL,
	[SID] [int] NULL,
	[GUID] [nvarchar](50) NULL,
	[TITLE] [nvarchar](500) NOT NULL,
	[NOTE] [nvarchar](max) NULL,
	[NOTE1] [nvarchar](500) NULL,
	[NOTE2] [nvarchar](50) NULL,
	[NOTE3] [nvarchar](50) NULL,
	[NOTE4] [nvarchar](50) NULL,
	[NOTE5] [nvarchar](50) NULL,
	[KEY_DATE] [datetime] NOT NULL,
	[KEY_MAN] [nvarchar](50) NOT NULL,
 CONSTRAINT [PK_SYSLOG] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]


