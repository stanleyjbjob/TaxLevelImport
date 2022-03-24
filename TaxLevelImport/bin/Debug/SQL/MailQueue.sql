CREATE TABLE [dbo].[MAILQUEUE](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[GUID] [nvarchar](50) NOT NULL,
	[FROM_ADDR] [nvarchar](50) NOT NULL,
	[FROM_NAME] [nvarchar](50) NULL,
	[TO_ADDR] [nvarchar](50) NOT NULL,
	[TO_NAME] [nvarchar](50) NULL,
	[SUBJECT] [nvarchar](500) NOT NULL,
	[BODY] [nvarchar](max) NOT NULL,
	[RETRY] [int] NOT NULL,
	[SUCCESS] [bit] NOT NULL,
	[SUSPEND] [bit] NOT NULL,
	[FREEZE_TIME] [datetime] NOT NULL,
	[KEY_DATE] [datetime] NOT NULL,
	[KEY_MAN] [nvarchar](50) NOT NULL,
	[NOTE] [nvarchar](50) NULL,
	[NOTE1] [nvarchar](50) NULL,
	[NOTE2] [nvarchar](50) NULL,
	[NOTE3] [nvarchar](50) NULL,
	[NOTE4] [nvarchar](50) NULL,
	[NOTE5] [nvarchar](50) NULL,
 CONSTRAINT [PK_MAILQUEUE] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]



