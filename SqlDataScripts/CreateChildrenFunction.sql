create FUNCTION dbo.CountChildren
(@ParentId varchar(16)) 
RETURNS bigint 
AS
BEGIN
declare @cChildren int
set @cChildren = 0

IF EXISTS (select ParentUniqueValue from dbo.tbCategoryRelation where ParentUniqueValue=@ParentId)
BEGIN 
     SELECT 
        @cChildren = Count(ParentUniqueValue) 
        FROM 
            dbo.tbCategoryRelation 
        WHERE 
            ParentUniqueValue = @ParentId
END 
  RETURN @cChildren 
END 