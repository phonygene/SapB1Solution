USE [MDR]
GO

INSERT INTO [dbo].[MGUIAPDetail_Import]
           ([ID]          --����MGUIAP_Import ID
           ,[LineId]      --�C
           ,[U_DocEntry]  -- AP�o�������s��
           ,[U_OBJTYPE]   -- �T�w��18
           ,[U_BELNR]     --AP�o���渹
           ,[U_BLDAT]     --���Ҥ��
           ,[U_VATDATE]   --�ӳ����
           ,[U_STCEG]     --�Τ@�s��
           ,[U_XBLNR]     --�o�����X(�r�y)
           ,[U_ZFORM_CODE]--21 �V �T�p���o�� 22 �V �G�p���o�� 25 �V �T�p�����Ⱦ��o��,�q�l�o�� 26 �V �J�`�n�J��դ��H�U�i���T�p��,�q�l�p��Τ@�o�� 27 �V �J�`�n�J��եH�U�i���G�p���o��,�����|�B����L�o�� 28 �V ��������
           ,[U_HWBAS]     --�o�����|�`���B
           ,[U_HWSTE]     --�o���|�B
           ,[U_TAX_TYPE]  --1-���| 2-�s�|3-�K�|
           ,[U_CUS_TYPE]  --0-�i���o�� 1-��L����-���t�|(�t�ȶO) 2-��L����-�~�t�|(���q�O)
           ,[U_AM_TYPE]   --1-�i����-�f�� 2-�i����-�T�w�겣 3-���i����-�f�� 4-���i����-�T�w�겣 5-�g�a6-�Ѳ��ѧQ
           ,[U_VATCODE]   --�T�w��100
           ,[U_BUKRS]     --�T�w��100
           ,[U_FA_DESC]   --�T��y�z
           ,[U_FA_QTY]    --�T��ƶq
           ,[U_FA_USE]	  --�T��γ~
		   ,[U_GatherMark]--�J�`���O �w�]�Ȭ� N
           ,[U_ConsolidQty]--�J�`�i��
           ,[U_MWSKZ]     --��~�|�N�X 21/22/25/26/27/28
           ,[U_LIFNR])    --�����ӥN�X(CardCode)
     VALUES
          (1, 1, 31, 18, '31', '2025-10-02 00:00:00.000', '2025-10-02 00:00:00.000',
        '53721864', 'RT20251004', '21', 200, 10, '1', '0', '1', '100', '100',
        '', 0,'', 'N', 0, '21', 'AA013')

GO


