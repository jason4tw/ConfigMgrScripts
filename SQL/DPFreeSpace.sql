SELECT 
	  SUBSTRING(SiteSystem, 13, PATINDEX('%\"%', SiteSystem) - 13) As [Server Name]
	  ,SUBSTRING(SiteObject, PATINDEX('%$%', SiteObject) - 1, 1) As [Drive Letter]
      ,[BytesTotal]  / 1024 / 1024 As [GB Total]
      ,[BytesFree] / 1024 / 1024 As [GB Free]
  FROM [CM_P01].[dbo].[v_SiteSystemSummarizer]
WHERE Role = 'SMS Distribution Point'
Order By [GB Free] Desc