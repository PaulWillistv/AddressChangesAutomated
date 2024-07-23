
drop table ChangesToApply


create table ChangesToApply(
RecordId int identity(1,1)
, LocationId varchar(200)
, ServingUnit varchar(200)
, Fttxdate	 varchar(200)
, FiberDistribution	 varchar(200)
, ChildProject	 varchar(200)
, ServingUnitNewName varchar(200)
, FttxSpeed	 varchar(200)
, CableModemSpeed	 varchar(200)
, ADSLSpeed	 varchar(200)
, TVType	 varchar(200)
, PhoneType	 varchar(200)
, Reason  varchar(900)
, SpreadsheetName   varchar(900)
, ChangeIsApplied     BIT DEFAULT 0
, DateOfData datetime default(getdate()) 
)

select * From tvcOperations.dbo.ChangesToApply

 