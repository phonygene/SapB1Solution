USE [MDR]
GO

INSERT INTO [dbo].[MGUIAP_Import]
           ([ID]          ---����MGUIAPDetail_Import ID
           ,[LineId]      --�PID��
           ,[CreateDate]  --�إߤ��
           ,[CreateBy]    --�إߪ�
           ,[UpdateDate]  --��s���
           ,[UpdateBy]    --��s��
           ,[U_DocEntry]  --AP�o�������s��
           ,[U_OBJTYPE]   --�T�w��18
           ,[DocNum]      --AP�o���渹
           ,[DocTotal]    --���|�`�p
           ,[VatSum])     --�o���|�`�p
     VALUES
		(1, 1, '2025-10-02 00:00:00.000', 'SOE', '2025-10-02 00:00:00.000','SOE',31 , '18', 31, 200, 10)
GO


