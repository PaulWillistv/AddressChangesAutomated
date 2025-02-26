﻿using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace AddressChangesAutomated
{
    internal class Program
    {
        /// <summary>
        ///  use tvcOperations
        /// go 
        /// select* from ChangesToApply order by recordid desc
        ///        create table ChangesToApply(
        ///       RecordId int identity(1,1)
        ///, LocationId varchar(200)
        ///, ServingUnit varchar(200)
        ///, Fttxdate varchar(200)
        ///, FiberDistribution varchar(200)
        ///, ChildProject varchar(200)
        ///, ServingUnitNewName varchar(200)
        ///, FttxSpeed varchar(200)
        ///, CableModemSpeed varchar(200)
        ///, ADSLSpeed varchar(200)
        ///, TVType varchar(200)
        ///, PhoneType varchar(200)
        ///, Reason varchar(900)
        ///, SpreadsheetName varchar(900)
        ///, ChangeIsApplied BIT DEFAULT 0
        ///, DateOfData datetime default(getdate()) 
        ///)
        /// </summary>
        /// <param name="args"></param>
        /// 

        static void queries()
        {
            /*
             --use AddressLoad

            select  CTA.StreetName   , aj.Streetname as 'aj.Streetname'
                , CTA.FTTxAvailableDate  , aj.FTTXDateAvailable as 'aj.FTTXDateAvailable'
	            , CTA.fttxspeed  , aj.fttxspeed as 'aj.fttxspeed'
	            , CTA.ServingUnit,  aj.servingunit   as 'aj.servingunit'
	            , CTA.[Distribution Fiber] , aj.DistFiber   as 'aj.DistFiber'
	            , CTA.CHILDProject  
	            , CTA.ProjectName 
            from  production360.OMNIA_ETVC_P_TVC_CM.dbo.addressjoin  aj   with (nolock)
            inner join  AddressLoad.dbo.ImportedSpreadsheet_AddressesPRESTAGING  CTA on CTA.TruvistaLocationID  = aj.srvlocation_locationid  
            where (
            isnull(CTA.StreetName,'') <> isnull(aj.Streetname ,'')
            or isnull(CTA.FTTxAvailableDate,'') <> isnull(aj.FTTXDateAvailable ,'')
            or isnull(CTA.FTTXSpeed,'') <> isnull(aj.fttxspeed ,'')
            or isnull(CTA.ServingUnit,'') <> isnull(aj.ServingUnit ,'')
            or isnull(CTA.[Distribution Fiber],'') <> isnull(aj.DistFiber ,'')  )
 
            select 
              CTA.StreetName   , ast.Streetname as 'ast.Streetname'
                , CTA.FTTxAvailableDate  , ast.FTTxAvailableDate as 'ast.FTTXDateAvailable'
	            , CTA.fttxspeed  , ast.fttxspeed as 'ast.fttxspeed'
	            , CTA.ServingUnit,  ast.servingunit   as 'ast.servingunit'
	            , CTA.[Distribution Fiber] , ast.Distribution   as 'ast.DistFiber'
	            , CTA.CHILDProject  
	            , CTA.ProjectName , ast.ProjectName 
             from AddressLoad.dbo.AddressStaging ast
            inner join  AddressLoad.dbo.ImportedSpreadsheet_AddressesPRESTAGING  CTA on CTA.TruvistaLocationID  = ast.TruvistaLocationID  
            where (
            isnull(CTA.StreetName,'') <> isnull(ast.Streetname ,'')
            or isnull(CTA.FTTxAvailableDate,'') <> isnull( ast.FTTxAvailableDate ,'')
            or isnull(CTA.FTTXSpeed,'') <> isnull( ast.fttxspeed ,'')
            or isnull(CTA.ServingUnit,'') <> isnull( ast.ServingUnit ,'')
            --or isnull(CTA.[Distribution Fiber],'') <> isnull( ast.Distribution ,'')  
            )
                        */
        }
        public struct ColumnNameAndPositon
        {
            public ColumnNameAndPositon(int colPosition, string columnName, bool hasColumnName, bool hasRecordsInColumn)
            {
                ColPosition = colPosition;
                ColumnName = columnName;
                HasColumnName = hasColumnName;
                HasRecordsInColumn = hasRecordsInColumn;
            }

            public int ColPosition { get; }
            public string ColumnName { get; }
            public bool HasColumnName { get; }
            public bool HasRecordsInColumn { get; }

        }

        static void Main(string[] args)
        {


            Console.WriteLine("Username:");
            string Ans_Username = Console.ReadLine();

            Console.WriteLine("Password:");
            string Ans_Password = Console.ReadLine();



            string path = @"\\ctcfile\public\AddressLoad\Changes\WIP\";

            Console.WriteLine(String.Format("Using path: {0}", path));
            Console.WriteLine(String.Format("Inserting into table: {0} in {1}", "ChangesToApply", "tvcOperations"));


            var directory = new DirectoryInfo(path);

            var myFile = (from f in directory.GetFiles()
                          orderby f.LastWriteTime descending
                          select f);

            foreach (FileInfo f in myFile)
            {

                if (IsInDatabase(f.Name))
                {
                    continue;
                }

                Console.WriteLine(f.Name);
                if (f.Name.StartsWith("~$"))
                {
                    continue;
                }

                if (f.Name.StartsWith("Changes_20230314.xlsx"))
                {
                    continue;
                }
                if (f.Name.StartsWith("Changes_202402.xlsx"))
                {
                    continue;
                }
                if (f.Name.ToUpper().Contains("IMPORTED"))
                {
                    continue;
                }


                int aa = BLL.GetRowCount(f.Name, "fttxdate", Ans_Username, Ans_Password);


                string UserProvidedExcelFileWithFullPath = "";
                UserProvidedExcelFileWithFullPath = path + "\\" + f.ToString();

                string SpreadsheetName = UserProvidedExcelFileWithFullPath;

                bool hdr7 = true;
                System.Data.DataSet DataSetFromExcel = Util.LoadExcelIntoDataTable(SpreadsheetName, hdr7, false, "Sheet1");
                int RecordCountInSpreadsheet = DataSetFromExcel.Tables[0].Rows.Count;


                System.Data.DataTable TableFetched = DataSetFromExcel.Tables[0];

                TableFetched.Columns[2].ColumnName = "Reason";
                TableFetched.AcceptChanges();


                ColumnNameAndPositon ColAndPos_LocationID = FirstColumnIsCorrectLabel(DataSetFromExcel, "LOCATIONID");
                ColumnNameAndPositon ColAndPos_ServingUnit = FirstColumnIsCorrectLabel(DataSetFromExcel, "SERVINGUNIT");
                ColumnNameAndPositon ColAndPos_Reason = FirstColumnIsCorrectLabel(DataSetFromExcel, "REASON");
                ColumnNameAndPositon ColAndPos_Fttxdate = FirstColumnIsCorrectLabel(DataSetFromExcel, "FTTXDATE");
                ColumnNameAndPositon ColAndPos_Funded = FirstColumnIsCorrectLabel(DataSetFromExcel, "FUNDED");
                ColumnNameAndPositon ColAndPos_FinanceRegion = FirstColumnIsCorrectLabel(DataSetFromExcel, "FINANCEREGION");
                ColumnNameAndPositon ColAndPos_FiberDistribution = FirstColumnIsCorrectLabel(DataSetFromExcel, "FIBERDISTRIBUTION");
                ColumnNameAndPositon ColAndPos_ChildProject = FirstColumnIsCorrectLabel(DataSetFromExcel, "CHILDPROJECT");
                ColumnNameAndPositon ColAndPos_ServingUnitNewName = FirstColumnIsCorrectLabel(DataSetFromExcel, "SERVINGUNITNEWNAME");
                ColumnNameAndPositon ColAndPos_InternetType = FirstColumnIsCorrectLabel(DataSetFromExcel, "INTERNETTYPE");

                if (ColAndPos_LocationID.HasRecordsInColumn && ColAndPos_Fttxdate.HasRecordsInColumn)
                {
                    DoFTTXDateUpdate(DataSetFromExcel, f.Name, Ans_Username, Ans_Password);
                }
                else if (ColAndPos_LocationID.HasRecordsInColumn && ColAndPos_FinanceRegion.HasRecordsInColumn)
                {
                    DoFinanceRegionUpdate(DataSetFromExcel, f.Name, Ans_Username, Ans_Password);
                }

                else if (ColAndPos_LocationID.HasRecordsInColumn && ColAndPos_Funded.HasRecordsInColumn)
                {
                    DoFundedUpdate(DataSetFromExcel, f.Name, Ans_Username, Ans_Password);
                }

                else if (ColAndPos_LocationID.HasRecordsInColumn && ColAndPos_FiberDistribution.HasRecordsInColumn)
                {
                    DoFiberDistributionUpdate(DataSetFromExcel, f.Name, Ans_Username, Ans_Password);

                    //foreach (DataRow row in DataSetFromExcel.Tables[0].Rows)
                    //{
                    //    string sql = String.Format(@"insert into ChangesToApply
                    //            ( 
                    //                LocationID
                    //                , FiberDistribution
                    //                , SpreadsheetName
                    //            ) 
                    //            values 
                    //            (
                    //                '{0}' 
                    //                ,'{1}' 
                    //                ,'{2}' 
                    //            )"
                    //            , row["LocationID"].ToString()
                    //            , row["FiberDistribution"].ToString() 
                    //            , f.Name);
                    //    DAL.RunNonQuery(sql);
                    //}

                }
                else if (ColAndPos_LocationID.HasRecordsInColumn && ColAndPos_ServingUnitNewName.HasRecordsInColumn)
                {
                    DoServingUnitNewNameUpdate(DataSetFromExcel, f.Name, Ans_Username, Ans_Password);
                }

                else if (ColAndPos_LocationID.HasRecordsInColumn && ColAndPos_InternetType.HasRecordsInColumn)
                {
                    DoInternetTypeUpdate(DataSetFromExcel, f.Name, Ans_Username, Ans_Password);
                }

                else if (ColAndPos_LocationID.HasRecordsInColumn && ColAndPos_ChildProject.HasRecordsInColumn)
                {
                    DoChildProjectUpdate(DataSetFromExcel, f.Name, Ans_Username, Ans_Password);
                }


                //else if (ColAndPos_LocationID.HasRecordsInColumn && ColAndPos_ChildProject.HasRecordsInColumn)
                //{
                //    DoChildProjectUpdate(DataSetFromExcel, f.Name);
                //}


                System.IO.File.Move(path + f.Name, path + f.Name + "_Imported.xlsx");


                string Alltext = @" select* from tvcoperations.dbo.ChangesToApply where SpreadsheetName like '" + f.Name + "' order by recordid desc";

                // Notepad.NotepadHelper.ShowMessage(Alltext.ToString(), "SQL script - Alltext");
            }
        }

        private static bool IsInDatabase(string name)
        {
            bool res = false;

            return res;
        }

        private static void DoInternetTypeUpdate(System.Data.DataSet DataSetFromExcel, string FileName, string Username, string Password)
        {
            foreach (DataRow row in DataSetFromExcel.Tables[0].Rows)
            {
                string sql = String.Format(@"insert into ChangesToApply
                                ( 
                                    LocationID
                                    , InternetType 
                                    , Reason
                                    , SpreadsheetName
                                ) 
                                values 
                                (
                                    '{0}' 
                                    ,'{1}' 
                                    ,'{2}'
                                    ,'{3}' 

                                )"
                        , row["LocationID"].ToString()
                        , row["InternetType"].ToString()
                        , row["Reason"].ToString()
                        , FileName);
                DAL.RunNonQuery(sql, Username, Password);
            }
        }

        private static void DoServingUnitNewNameUpdate(System.Data.DataSet DataSetFromExcel, string FileName, string Username, string Password)
        {
            /* 
            update aj set aj.ServingUnit = cta.ServingUnitNewName 
            from  OMNIA_ETVC_P_TVC_CM.dbo.addressjoin aj
            inner join tvcOperations.dbo.ChangesToApply  CTA on CTA.locationid = aj.srvlocation_locationid   
            where cta.SpreadsheetName  = 'Changes_Carter PON multiple changes_20240327 - ServingUnitNewName.xlsx'

            update sl set sl.new_ServingUnit = cta.ServingUnitNewName  
            from TVCO360_MSCRM.dbo.chr_servicelocation sl 
            inner join tvcOperations.dbo.ChangesToApply  CTA on CTA.locationid collate database_default=sl.chr_masterlocationid collate database_default
            where cta.SpreadsheetName  = 'Changes_Carter PON multiple changes_20240327 - ServingUnitNewName.xlsx'

            update cta set cta.changeisapplied = 1 from tvcOperations.dbo.ChangesToApply  CTA where cta.SpreadsheetName  = 'Changes_Carter PON multiple changes_20240327 - ServingUnitNewName.xlsx' 
            */

            StringBuilder sb = new StringBuilder();
            sb.Append(" /*   " + Environment.NewLine);
            sb.Append("   use tvcOperations " + Environment.NewLine);

            sb.Append("update  aj set  aj.ServingUnit = cta.ServingUnitNewName " + Environment.NewLine);
            sb.Append("-- select  aj.ServingUnit , CTA.ServingUnitNewName" + Environment.NewLine);
            sb.Append("from  OMNIA_ETVC_P_TVC_CM.dbo.addressjoin aj" + Environment.NewLine);
            sb.Append("inner join ( select locationid ,ServingUnitNewName ,spreadsheetname  from ChangesToApply where locationid is not null and ServingUnitNewName is not null )  " + Environment.NewLine);
            sb.Append("    CTA on CTA.locationid collate database_default =  aj.SrvLocation_LocationID collate database_default" + Environment.NewLine);
            sb.Append(" where cta.spreadsheetname ='" + FileName + @"'  " + Environment.NewLine);

            sb.Append("" + Environment.NewLine);


            sb.Append("update sl set sl.new_ServingUnit = CTA.ServingUnitNewName" + Environment.NewLine);
            sb.Append("-- select   sl.new_ServingUnit , CTA.ServingUnitNewName" + Environment.NewLine);
            sb.Append("from TVCO360_MSCRM.dbo.chr_servicelocation sl" + Environment.NewLine);
            sb.Append("inner join ( select locationid ,ServingUnitNewName ,spreadsheetname  from ChangesToApply where locationid is not null and ServingUnitNewName is not null )  " + Environment.NewLine);
            sb.Append("    CTA on CTA.locationid collate database_default = sl.chr_masterlocationid collate database_default" + Environment.NewLine);
            sb.Append(" where cta.spreadsheetname ='" + FileName + @"'  " + Environment.NewLine);

            sb.Append("" + Environment.NewLine);

            sb.Append(" " + Environment.NewLine);
            sb.Append(" */" + Environment.NewLine);

            Notepad.NotepadHelper.ShowMessage(sb.ToString(), "SQL script - Alltext");

            foreach (DataRow row in DataSetFromExcel.Tables[0].Rows)
            {
                string sql = String.Format(@"insert into ChangesToApply
                                ( 
                                    LocationID
                                    , ServingUnitNewName 
                                    , Reason
                                    , SpreadsheetName
                                ) 
                                values 
                                (
                                    '{0}' 
                                    ,'{1}' 
                                    ,'{2}'
                                    ,'{3}' 

                                )"
                        , row["LocationID"].ToString()
                        , row["ServingUnitNewName"].ToString()
                        , row["Reason"].ToString()
                        , FileName);
                DAL.RunNonQuery(sql, Username, Password);
            }
        }
        private static void DoChildProjectUpdate(System.Data.DataSet DataSetFromExcel, string FileName, string Username, string Password)
        {
            /*
            update sl set sl.new_ChildProjectName = cta.ChildProject 
            from TVCO360_MSCRM.dbo.chr_servicelocation sl 
            inner join tvcOperations.dbo.ChangesToApply  CTA on CTA.locationid collate database_default=sl.chr_masterlocationid collate database_default
            where cta.SpreadsheetName  = 'Changes_Carter PON multiple changes_20240327 - childproject.xlsx'

            update cta set cta.changeisapplied = 1 from tvcOperations.dbo.ChangesToApply  CTA where cta.SpreadsheetName  = 'Changes_Carter PON multiple changes_20240327 - childproject.xlsx'
            */
            StringBuilder sb = new StringBuilder();
            sb.Append(" /*   " + Environment.NewLine);
            sb.Append("   use tvcOperations " + Environment.NewLine);



            sb.Append("update sl set sl.new_ChildProjectName = CTA.ChildProject" + Environment.NewLine);
            sb.Append("-- select   sl.new_ChildProjectName , CTA.ChildProject" + Environment.NewLine);
            sb.Append("from TVCO360_MSCRM.dbo.chr_servicelocation sl" + Environment.NewLine);
            sb.Append("inner join ( select locationid ,ChildProject,spreadsheetname  from ChangesToApply where locationid is not null and ChildProject is not null ) " + Environment.NewLine);

            sb.Append("    CTA on CTA.locationid collate database_default = sl.chr_masterlocationid collate database_default" + Environment.NewLine);
            sb.Append("where cta.spreadsheetname ='" + FileName + @"'  " + Environment.NewLine);


            sb.Append(" " + Environment.NewLine);
            sb.Append(" */" + Environment.NewLine);

            Notepad.NotepadHelper.ShowMessage(sb.ToString(), "SQL script - Alltext");



            foreach (DataRow row in DataSetFromExcel.Tables[0].Rows)
            {
                string sql = String.Format(@"insert into ChangesToApply
                                ( 
                                    LocationID
                                    , ChildProject
                                    , Reason
                                    , SpreadsheetName
                                ) 
                                values 
                                (
                                    '{0}' 
                                    ,'{1}' 
                                    ,'{2}'
                                    ,'{3}' 

                                )"
                        , row["LocationID"].ToString()
                        , row["ChildProject"].ToString()
                        , row["Reason"].ToString()
                        , FileName);
                DAL.RunNonQuery(sql, Username, Password);
            }

        }


        private static void DoFiberDistributionUpdate(System.Data.DataSet DataSetFromExcel, string FileName, string Username, string Password)
        {
            /* 
update cta set cta.changeisapplied = 1 from tvcOperations.dbo.ChangesToApply  CTA where cta.SpreadsheetName  = 'Changes_SopertonOverbuildFDC_20240319 -fttxdate.xlsx'

                select  cta.FiberDistribution, aj.DistFiber , * 
                from  OMNIA_ETVC_P_TVC_CM.dbo.addressjoin aj
                inner join (
			                select locationid ,FiberDistribution from tvcOperations.dbo.ChangesToApply 
			                where locationid is not null 
			                and FiberDistribution  is not null
			                and  SpreadsheetName = 'Changes_SopertonOverbuildFDC_20240315.xlsx'
		                ) 
	                CTA on CTA.locationid = aj.srvlocation_locationid  
                where isnull( cta.FiberDistribution ,'')  <> isnull( aj.DistFiber ,'') 
                */

            /*
             
            update sl set sl.new_DistFiber = cta.FiberDistribution 
            from TVCO360_MSCRM.dbo.chr_servicelocation sl 
            inner join tvcOperations.dbo.ChangesToApply  CTA on CTA.locationid collate database_default=sl.chr_masterlocationid collate database_default
            where cta.SpreadsheetName  =  'Changes_NGA Stphns-Dortch FTTx update fiber distro 20240328.xlsx'
            */

            /*
             
            update aj set aj.DistFiber = cta.FiberDistribution 
            from  OMNIA_ETVC_P_TVC_CM.dbo.addressjoin aj
            inner join tvcOperations.dbo.ChangesToApply  CTA on CTA.locationid = aj.srvlocation_locationid   
            where cta.SpreadsheetName  = 'Changes_NGA Stphns-Dortch FTTx update fiber distro 20240328.xlsx'
                        */


            /*
             update aj set  aj.DistFiber = cta.FiberDistribution 
            from  OMNIA_ETVC_P_TVC_CM.dbo.addressjoin aj
            inner join (
			            select locationid ,FiberDistribution from tvcOperations.dbo.ChangesToApply 
			            where locationid is not null 
			            and FiberDistribution  is not null
			            and  SpreadsheetName = 'Changes_SopertonOverbuildFDC_20240315.xlsx'
		            ) 
	            CTA on CTA.locationid = aj.srvlocation_locationid  
            where isnull( cta.FiberDistribution ,'')  <> isnull( aj.DistFiber ,'') */


            StringBuilder sb = new StringBuilder();
            sb.Append(" /*   " + Environment.NewLine);
            sb.Append("   use tvcOperations " + Environment.NewLine);

            sb.Append("update  aj set  aj.DistFiber = cta.FiberDistribution " + Environment.NewLine);
            sb.Append("-- select  aj.DistFiber , CTA.FiberDistribution" + Environment.NewLine);
            sb.Append("from  OMNIA_ETVC_P_TVC_CM.dbo.addressjoin aj" + Environment.NewLine);
            sb.Append("inner join ( select locationid ,FiberDistribution ,spreadsheetname  from ChangesToApply where locationid is not null and FiberDistribution is not null )  " + Environment.NewLine);
            sb.Append("    CTA on CTA.locationid collate database_default =  aj.SrvLocation_LocationID collate database_default" + Environment.NewLine);
            sb.Append(" where cta.spreadsheetname ='" + FileName + @"'  " + Environment.NewLine);

            sb.Append("" + Environment.NewLine);


            sb.Append("update sl set sl.new_DistFiber = CTA.FiberDistribution" + Environment.NewLine);
            sb.Append("-- select   sl.new_DistFiber , CTA.FiberDistribution" + Environment.NewLine);
            sb.Append("from TVCO360_MSCRM.dbo.chr_servicelocation sl" + Environment.NewLine);
            sb.Append("inner join ( select locationid ,FiberDistribution ,spreadsheetname  from ChangesToApply where locationid is not null and FiberDistribution is not null )  " + Environment.NewLine);
            sb.Append("    CTA on CTA.locationid collate database_default = sl.chr_masterlocationid collate database_default" + Environment.NewLine);
            sb.Append(" where cta.spreadsheetname ='" + FileName + @"'  " + Environment.NewLine);

            sb.Append("" + Environment.NewLine);

            sb.Append(" " + Environment.NewLine);
            sb.Append(" */" + Environment.NewLine);

            Notepad.NotepadHelper.ShowMessage(sb.ToString(), "SQL script - Alltext");


            foreach (DataRow row in DataSetFromExcel.Tables[0].Rows)
            {
                string sql = String.Format(@"insert into ChangesToApply
                                ( 
                                    LocationID
                                    , FiberDistribution
                                    , Reason
                                    , SpreadsheetName
                                ) 
                                values 
                                (
                                    '{0}' 
                                    ,'{1}' 
                                    ,'{2}'
                                    ,'{3}' 

                                )"
                        , row["LocationID"].ToString()
                        , row["FiberDistribution"].ToString()
                        , row["Reason"].ToString()
                        , FileName);
                DAL.RunNonQuery(sql, Username, Password);
            }

        }


        private static void DoFundedUpdate(System.Data.DataSet DataSetFromExcel, string FileName, string Username, string Password)
        {

            StringBuilder sb = new StringBuilder();
            sb.Append(" /*   " + Environment.NewLine);
            sb.Append("   use tvcOperations " + Environment.NewLine);


            sb.Append("update sl set sl.new_Funded = CTA.funded" + Environment.NewLine);
            sb.Append("-- select   sl.new_Funded , CTA.funded" + Environment.NewLine);
            sb.Append("from TVCO360_MSCRM.dbo.chr_servicelocation sl" + Environment.NewLine);
            sb.Append("inner join ( select locationid ,funded = case when funded = 'Unfunded' then 'U' when funded ='Funded' then 'F' else '' end   ,spreadsheetname  from ChangesToApply where locationid is not null and funded is not null )  " + Environment.NewLine);
            sb.Append("    CTA on CTA.locationid collate database_default = sl.chr_masterlocationid collate database_default" + Environment.NewLine);
            sb.Append(" where cta.spreadsheetname ='" + FileName + @"'  " + Environment.NewLine);

            sb.Append("" + Environment.NewLine);

            sb.Append(" " + Environment.NewLine);
            sb.Append(" */" + Environment.NewLine);

            Notepad.NotepadHelper.ShowMessage(sb.ToString(), "SQL script - Alltext");


            foreach (DataRow row in DataSetFromExcel.Tables[0].Rows)
            {
                string sql = String.Format(@"insert into ChangesToApply
                                ( 
                                    LocationID
                                    , Funded
                                    , Reason
                                    , SpreadsheetName
                                ) 
                                values 
                                (
                                    '{0}' 
                                    ,'{1}' 
                                    ,'{2}'
                                    ,'{3}' 

                                )"
                        , row["LocationID"].ToString()
                        , row["Funded"].ToString()
                        , row["Reason"].ToString()
                        , FileName);



                DAL.RunNonQuery(sql, Username, Password);
            }

        }

        private static void DoFinanceRegionUpdate(System.Data.DataSet DataSetFromExcel, string FileName, string Username, string Password)
        {

            StringBuilder sb = new StringBuilder();
            sb.Append(" /*   " + Environment.NewLine);
            sb.Append("   use tvcOperations " + Environment.NewLine);


            sb.Append("update sl set sl.new_FinanceRegion = CTA.FinanceRegion" + Environment.NewLine);
            sb.Append("-- select   sl.new_FinanceRegion , CTA.FinanceRegion" + Environment.NewLine);
            sb.Append("from TVCO360_MSCRM.dbo.chr_servicelocation sl" + Environment.NewLine);
            sb.Append("inner join ( select locationid ,FinanceRegion,spreadsheetname  from ChangesToApply where locationid is not null and FinanceRegion is not null ) " + Environment.NewLine);

            sb.Append("    CTA on CTA.locationid collate database_default = sl.chr_masterlocationid collate database_default" + Environment.NewLine);
            sb.Append("where cta.spreadsheetname ='" + FileName + @"'  " + Environment.NewLine);


            sb.Append(" " + Environment.NewLine);
            sb.Append(" */" + Environment.NewLine);

            Notepad.NotepadHelper.ShowMessage(sb.ToString(), "SQL script - Alltext");


            foreach (DataRow row in DataSetFromExcel.Tables[0].Rows)
            {
                string sql = String.Format(@"insert into ChangesToApply
                                ( 
                                    LocationID
                                    , FinanceRegion
                                    , Reason
                                    , SpreadsheetName
                                ) 
                                values 
                                (
                                    '{0}' 
                                    ,'{1}' 
                                    ,'{2}'
                                    ,'{3}' 

                                )"
                        , row["LocationID"].ToString()
                        , row["FinanceRegion"].ToString()
                        , row["Reason"].ToString()
                        , FileName);



                DAL.RunNonQuery(sql, Username, Password);
            }

        }



        private static void DoFTTXDateUpdate(System.Data.DataSet DataSetFromExcel, string FileName, string Username, string Password)
        {

            StringBuilder sb = new StringBuilder();
            sb.Append(" /*   " + Environment.NewLine);
            sb.Append("   use tvcOperations " + Environment.NewLine);


            sb.Append("update aj set aj.FTTXDateAvailable = CTA.fttxdate" + Environment.NewLine);
            sb.Append("-- select  aj.FTTXDateAvailable , CTA.fttxdate" + Environment.NewLine);
            sb.Append("from  OMNIA_ETVC_P_TVC_CM.dbo.addressjoin aj" + Environment.NewLine);
            sb.Append("inner join ( select locationid ,fttxdate ,spreadsheetname  from ChangesToApply where locationid is not null and fttxdate is not null )  CTA on CTA.locationid = aj.srvlocation_locationid " + Environment.NewLine);
            sb.Append("where cta.spreadsheetname = '" + FileName + @"' " + Environment.NewLine);

            sb.Append("" + Environment.NewLine);

            sb.Append("update sl set sl.new_FTTXDateAvailable = CTA.fttxdate" + Environment.NewLine);
            sb.Append("-- select   sl.new_FTTXDateAvailable , CTA.fttxdate" + Environment.NewLine);
            sb.Append("from TVCO360_MSCRM.dbo.chr_servicelocation sl" + Environment.NewLine);
            sb.Append("inner join ( select locationid ,fttxdate,spreadsheetname  from ChangesToApply where locationid is not null and fttxdate is not null ) " + Environment.NewLine);

            sb.Append("    CTA on CTA.locationid collate database_default = sl.chr_masterlocationid collate database_default" + Environment.NewLine);
            sb.Append("where cta.spreadsheetname ='" + FileName + @"'  " + Environment.NewLine);


            sb.Append(" " + Environment.NewLine);
            sb.Append(" */" + Environment.NewLine);


            /* select cta.FiberDistribution 
	                , CTA.fttxdate 
	                , CTA.fttxspeed 
	                , aj.FTTXDateAvailable
	                , aj.DistFiber 
	                , sl.new_FTTXDateAvailable 
	                , sl.new_FTTXSpeed 
	                , sl.new_DistFiber
                from  OMNIA_ETVC_P_TVC_CM.dbo.addressjoin aj
                inner join tvcOperations.dbo.ChangesToApply  CTA on CTA.locationid = aj.srvlocation_locationid  
                left outer join TVCO360_MSCRM.dbo.chr_servicelocation sl on sl.chr_masterlocationid = aj.srvlocation_locationid  
                where cta.SpreadsheetName  = 'Changes_SopertonOverbuildFDC_20240315.xlsx' /*fibdist*/
            //         where cta.SpreadsheetName = 'Changes_Soperton Overbuild Fttx date 3.14.2024.xlsx' /*fttxdate*/*/


            Notepad.NotepadHelper.ShowMessage(sb.ToString(), "SQL script - Alltext");


            foreach (DataRow row in DataSetFromExcel.Tables[0].Rows)
            {
                string sql = String.Format(@"insert into ChangesToApply
                                ( 
                                    LocationID
                                    , fttxdate
                                    , Reason
                                    , SpreadsheetName
                                ) 
                                values 
                                (
                                    '{0}' 
                                    ,'{1}' 
                                    ,'{2}'
                                    ,'{3}' 

                                )"
                        , row["LocationID"].ToString()
                        , row["Fttxdate"].ToString()
                        , row["Reason"].ToString()
                        , FileName);



                DAL.RunNonQuery(sql, Username, Password);
            }

        }

        private static ColumnNameAndPositon FirstColumnIsCorrectLabel(DataSet dataSetFromExcel, string ColumnNameToFind)
        {

            System.Data.DataTable dtNames = dataSetFromExcel.Tables["MainTable"];

            int ColPosition = 0;
            bool HasColumnName = false;
            string ColumnName = "";
            bool HasRecordsInColumn = false;
            int Count = 0;

            foreach (DataColumn dc in dtNames.Columns)
            {
                if (dc.ColumnName.ToUpper().Trim() == ColumnNameToFind.ToUpper().Trim())
                {
                    HasColumnName = true;
                    ColPosition = Count;
                    ColumnName = dc.ColumnName.ToUpper().Trim();


                    int CheckCount = 0;
                    foreach (DataRow row in dtNames.Rows)
                    {
                        if (CheckCount > 2)
                        {
                            break;
                        }
                        if (row[dc.ColumnName].ToString().Length > 0)
                        {
                            HasRecordsInColumn = true;
                        }
                        CheckCount = CheckCount + 1;
                    }




                    break;
                }
                Count = Count + 1;
            }

            ColumnNameAndPositon res = new ColumnNameAndPositon(ColPosition, ColumnNameToFind, HasColumnName, HasRecordsInColumn);
            return res;
        }
    }
}
