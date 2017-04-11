USE [msdb]
GO

/****** Object:  Job [Purity_PowerBI_updates]    Script Date: 04/11/2017 13:37:21 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** Object:  JobCategory [[Uncategorized (Local)]]]    Script Date: 04/11/2017 13:37:22 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'Purity_PowerBI_updates', 
		@enabled=1, 
		@notify_level_eventlog=0, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'No description available.', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** Object:  Step [CopyProdCounts]    Script Date: 04/11/2017 13:37:22 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'CopyProdCounts', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=3, 
		@on_success_step_id=0, 
		@on_fail_action=3, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'SSIS', 
		@command=N'/FILE "C:\Users\administrator.PURITY\Documents\Integration Services Script Task\GetProdCounts.dtsx" /CHECKPOINTING OFF /REPORTING E', 
		@database_name=N'master', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** Object:  Step [All steps]    Script Date: 04/11/2017 13:37:22 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'All steps', 
		@step_id=2, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'/*Replace Ship header table*/

DROP TABLE [Purity_PowerBI_data].[dbo].[Header_Data]

SELECT [BillAddr2]
      ,[CustID]
      ,[DateCancelled]
      ,[OrdDate]
      ,[OrdNbr]
      ,[ShipName]
      ,[ShipperID]
      ,[SOTypeID]
      ,[Zone]
      ,[CpnyID]
      ,[InvcDate]
      ,[User7] As [Route]
INTO [Purity_PowerBI_data].[dbo].[Header_Data]
FROM         [PurityApp].[dbo].[SOShipHeader] 
WHERE     ([InvcDate] > ''2014-09-30'')

/*Replace Ship Line Data table*/
DROP TABLE [Purity_PowerBI_data].[dbo].[Line_Data]

SELECT [CnvFact]
	 , [Descr]
	 , [InvtID]
	 , [PurityApp].[dbo].[SOShipLine].[OrdNbr]
	 , [OrigInvtID]
	 , [OrigShipperID]
	 , [QtyOrd]
	 , [QtyPick]
	 , [QtyPrevShip]
	 , [QtyShip]
	 , [PurityApp].[dbo].[SOShipLine].[ShipperID]
	 , [SiteID]
	 , [SlsPrice]
	 , [UnitDesc]
	 , [UnitMultDiv]
	 , [PurityApp].[dbo].[SOShipLine].[CpnyID]
	 , [CuryTotMerch]
INTO [Purity_PowerBI_data].[dbo].[Line_Data]
FROM [PurityApp].[dbo].[SOShipLine] WITH (INDEX (SOShipLine2)) INNER JOIN
     [Purity_PowerBI_data].[dbo].[Header_Data] ON [Purity_PowerBI_data].[dbo].[Header_Data].[ShipperID] = [PurityApp].[dbo].[SOShipLine].[ShipperID]
WHERE     ([Purity_PowerBI_data].[dbo].[Header_Data].[InvcDate] > ''2014-09-30'')

/*Replace Customer table*/
DROP TABLE [Purity_PowerBI_data].[dbo].[Customer]

SELECT * 
INTO [Purity_PowerBI_data].[dbo].[Customer]
FROM [PurityApp].[dbo].[Customer]

/*Replace SO Type table*/
DROP TABLE [Purity_PowerBI_data].[dbo].[SOType]

SELECT * 
INTO [Purity_PowerBI_data].[dbo].[SOType]
FROM [PurityApp].[dbo].[SOType]

/*Replace Customer EDI table*/
DROP TABLE [Purity_PowerBI_data].[dbo].[CustomerEDI]

SELECT *
INTO [Purity_PowerBI_data].[dbo].[CustomerEDI]
FROM [PurityApp].[dbo].[CustomerEDI]

/*Replace Inventory table*/
DROP TABLE [Purity_PowerBI_data].[dbo].[Inventory]

SELECT *
INTO [Purity_PowerBI_data].[dbo].[Inventory]
FROM [PurityApp].[dbo].[Inventory]

DROP TABLE [Purity_PowerBI_data].[dbo].[vr]

SELECT * 
INTO [Purity_PowerBI_data].[dbo].[vr]
FROM [PurityApp].[dbo].[vr_08650s]

/*Replace Product Class table*/
DROP TABLE [Purity_PowerBI_data].[dbo].[ProductClass]

SELECT *
INTO [Purity_PowerBI_data].[dbo].[ProductClass]
FROM [PurityApp].[dbo].[ProductClass]

/*Replace Product Class table*/
DROP TABLE [Purity_PowerBI_data].[dbo].[Header_Data_All_Items2]

SELECT [BillAddr2]
      ,[BuildCmpltDate]
      ,[CancelBO]
      ,[CustID]
      ,[DateCancelled]
      ,[OrdDate]
      ,[OrdNbr]
      ,[ShipName]
      ,[ShipperID]
      ,[SOTypeID]
      ,[Zone]
      ,[CpnyID]
      ,[InvcDate]
      ,[User7] As [Route]
INTO [Purity_PowerBI_data].[dbo].[Header_Data_All_Items2]
FROM         [PurityApp].[dbo].[SOShipHeader]
WHERE ([InvcDate]>''2015-09-30'' or ([InvcDate] = ''1900-01-01 00:00:00'' and [OrdDate] > ''2015-09-30''))


DROP TABLE [Purity_PowerBI_data].[dbo].[Line_Data_All_Items4]

SELECT [CnvFact]
      ,[Descr]
      ,[InvtID]
      ,[OrdNbr]
      ,[OrigInvtID]
      ,[OrigShipperID]
      ,[QtyOrd]
      ,[QtyPick]
      ,[QtyPrevShip]
      ,[QtyShip]
      ,[ShipperID]
      ,[SiteID]
      ,[SlsPrice]
      ,[UnitDesc]
      ,[UnitMultDiv]
      ,[CpnyID]
      ,[CuryTotMerch]
  INTO [Purity_PowerBI_data].[dbo].[Line_Data_All_Items4]
  
  FROM [PurityApp].[dbo].[SOShipLine] 
  where [PurityApp].[dbo].[SOShipLine].[ShipperID] in (select shipperid from [Purity_PowerBI_data].[dbo].[Header_Data_All_Items2])


', 
		@database_name=N'PurityApp', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'Load previous day''s data', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20160727, 
		@active_end_date=99991231, 
		@active_start_time=40000, 
		@active_end_time=235959, 
		@schedule_uid=N'3941cc9e-090a-4e29-af98-4d39c0c1c580'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:

GO


