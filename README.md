# Office 365 Group Creator

## Description
Creates Unified Groups in Office 365 based on database input.

## Dependencies
An account in Office 365 that can manage groups. An SQL Server database (see DDL below).

## Database objects
```sql
CREATE TABLE dbo.Office365Group(
	id int IDENTITY NOT NULL,
	groupName nvarchar(200) NOT NULL,
	groupOwner nvarchar(200) NOT NULL,
	groupSuffix nvarchar(50) NULL,
	namingScheme nvarchar(100) NOT NULL,
	groupGuid uniqueidentifier NULL,
	generatedDisplayName nvarchar(250) NULL,
	generatedAlias nvarchar(250) NULL,
	created datetime NULL,
	error nvarchar(1000) NULL,
    PRIMARY KEY (id)
);

CREATE PROCEDURE [dbo].[spO365GetGroupData]
AS
BEGIN
	SET NOCOUNT ON;

    SELECT
        id,
        groupName,
        groupOwner,
        groupSuffix,
        groupGuid,
        namingScheme,
        generatedAlias,
        generatedDisplayName,
        created,
        error
    FROM
        dbo.Office365Group
    WHERE
        created IS NULL
    AND
        error IS NULL;
END;

CREATE PROCEDURE [dbo].[spO365SetGroupData]
    @id int,
    @guid uniqueidentifier = NULL,
    @displayName nvarchar(250) = NULL,
    @alias nvarchar(250) = NULL,
    @created datetime = NULL,
    @error nvarchar(1000) = NULL
AS
BEGIN
	SET NOCOUNT ON;
    
    IF NOT EXISTS (SELECT 1 FROM dbo.Office365Group WHERE id = @id)
        THROW 50000, 'A group with requested ID was not found.', 1;

    UPDATE dbo.Office365Group SET
        groupGuid = ISNULL(@guid, groupGuid),
        generatedDisplayName = ISNULL(@displayName, generatedDisplayName),
        generatedAlias = ISNULL(@alias, generatedAlias),
        created = ISNULL(@created, created),
        error = ISNULL(@error, error)
    WHERE
        id = @id;
END;
```


