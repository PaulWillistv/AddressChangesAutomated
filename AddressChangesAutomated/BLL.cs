﻿using System.Data;

namespace AddressChangesAutomated
{
    internal class BLL
    {
        public static int GetRowCount(string FileName, string ColumnToCheck, string uname, string pwd)
        {
            /*
                select ' , ' +COLUMN_NAME+ '  is null'
                from INFORMATION_SCHEMA.columns where table_name = 'ChangesToApply'
                and COLUMN_NAME <> 'LocationId'
                and COLUMN_NAME <> 'recordid'
                and COLUMN_NAME <> 'DateOfData'
                and COLUMN_NAME <> 'ChangeIsApplied'
                and COLUMN_NAME <> 'SpreadsheetName'
                and COLUMN_NAME <> 'Reason'*/
            int res = 0;
            string sql = @"
                        select *  from ChangesToApply
                        where SpreadsheetName = '" + FileName + @"'
                        and " + ColumnToCheck + @" is not null
                        and servingunit is   null
                        and FiberDistribution  is   null
                        and ChildProject  is   null
                        and ServingUnitNewName   is   null
                        and FttxSpeed   is   null
                        and CableModemSpeed    is   null
                        and ADSLSpeed    is   null
                        and TVType    is   null
                        and TVType    is   null
                        and PhoneType    is   null";

            DataTable resDatatable = DAL.GetDataTable(sql, uname, pwd);
            res = resDatatable.Rows.Count;


            return res;
        }
    }
}
