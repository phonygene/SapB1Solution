USE [MDR]
GO

INSERT INTO [dbo].[MGUIAP_Import]
           ([ID]          ---對應MGUIAPDetail_Import ID
           ,[LineId]      --同ID號
           ,[CreateDate]  --建立日期
           ,[CreateBy]    --建立者
           ,[UpdateDate]  --更新日期
           ,[UpdateBy]    --更新者
           ,[U_DocEntry]  --AP發票內部編號
           ,[U_OBJTYPE]   --固定填18
           ,[DocNum]      --AP發票單號
           ,[DocTotal]    --未稅總計
           ,[VatSum])     --發票稅總計
     VALUES
		(1, 1, '2025-10-02 00:00:00.000', 'SOE', '2025-10-02 00:00:00.000','SOE',31 , '18', 31, 200, 10)
GO


