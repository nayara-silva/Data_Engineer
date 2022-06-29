SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[uspGeocoding]
@Address varchar(80) = NULL OUTPUT,
@City varchar(40) = NULL OUTPUT,
@State varchar(40) = NULL OUTPUT,
@Country varchar(40) = NULL OUTPUT,
@PostalCode varchar(20) = NULL OUTPUT,
@GPSLatitude numeric(9,6) = NULL OUTPUT,
@GPSLongitude numeric(9,6) = NULL OUTPUT,
@Key varchar(80) = NULL OUTPUT

AS
BEGIN
 SET NOCOUNT ON

DECLARE @URL varchar (MAX)
SET @URL = 'http://dev.virtualearth.net/REST/v1/Locations/' +
@Country + '/' +
CASE WHEN @State IS NOT NULL THEN @State ELSE '-' END + '/' +
CASE WHEN @City IS NOT NULL THEN @City ELSE '-' END + '/' +
CASE WHEN @PostalCode IS NOT NULL THEN @PostalCode ELSE '-' END + '/' +
CASE WHEN @Address IS NOT NULL THEN @Address ELSE '-' END + 
'?key=' + @Key

DECLARE @Response varchar(MAX)
DECLARE @XML xml
DECLARE @Obj int 
DECLARE @Result int 
DECLARE @HTTPStatus int 
DECLARE @ErrorMsg varchar(MAX)
DECLARE @ResponseText varchar(8000)
DECLARE @description varchar (300)

EXEC @Result = sp_OACreate 'MSXML2.ServerXMLHttp', @Obj OUT  

IF @Result <> 0 RAISERROR('Unable to open HTTP connection.', 10, 1);

BEGIN TRY
	EXEC @Result = sp_OAMethod @Obj, 'open', NULL, 'GET', @URL, false
	EXEC @Result = sp_OAMethod @Obj, 'setRequestHeader', NULL, 'Content-Type', 'application/x-www-form-urlencoded'
	EXEC @Result = sp_OAMethod @Obj, send, NULL, ''
	EXEC @Result = sp_OAGetProperty @Obj, 'status', @HTTPStatus OUT 
	EXEC @Result = sp_OAGetProperty @Obj, 'responseText', @ResponseText OUT
	--EXEC @Result = sp_OAGetProperty @Obj, 'responseXML.xml', @Response OUT 
	EXEC @Result = sp_OAGetErrorInfo @Obj, @Response OUT, @description OUT;
	PRINT @Response
 
END TRY
 BEGIN CATCH
 SET @ErrorMsg = ERROR_MESSAGE()
 END CATCH

 EXEC @Result = sp_OADestroy @Obj

IF (@ErrorMsg IS NOT NULL) OR (@HTTPStatus <> 200) BEGIN
 SET @ErrorMsg = 'Error in spGeocode: ' + ISNULL(@ErrorMsg, 'HTTP result is: ' + CAST(@HTTPStatus AS varchar(10)))
 RAISERROR(@ErrorMsg, 16, 1, @HTTPStatus)
 RETURN 
 END

SET @GPSLatitude = JSON_VALUE(@ResponseText, '$.resourceSets[0].resources[0].point.coordinates[0]') 
SET @GPSLongitude = JSON_VALUE(@ResponseText, '$.resourceSets[0].resources[0].point.coordinates[1]') 

  SELECT 
  @Country AS Country,
  @State AS State,
  @City AS City,
  @Address AS Address,
  @PostalCode AS [Postal Code],
  @GPSLatitude AS Latitude,
  @GPSLongitude AS Longitude

END
GO
