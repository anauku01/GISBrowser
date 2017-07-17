using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace COMPGIS1
{
    public class CompWorksDBConnect
    {

        public string _connectionstring = "";
        private string _dbfilename = "";
        private string _systemfilename = "";
        private string _companyname = "";
        private List<RiskRankingRange> _riskRankingRangesOverall = new List<RiskRankingRange>();
        private List<RiskRankingRange> _riskRankingRangesSusc = new List<RiskRankingRange>();
        private List<RiskRankingRange> _riskRankingRangesCons = new List<RiskRankingRange>();

        private Configuration _configuration;

        public const string CW_XRef_Fieldname = "MapFeatureID";

        public bool _isopen = false;
        public string CompanyName {
            set { }
            get { return _companyname; }
        }

        public List<RiskRankingRange> RiskRankingRanges
        {
            get { return _riskRankingRangesOverall; }
            set { _riskRankingRangesOverall = value; }
        }

        public List<RiskRankingRange> RiskRankingRangesSusc
        {
            get { return _riskRankingRangesSusc; }
            set { _riskRankingRangesSusc = value; }
        }

        public List<RiskRankingRange> RiskRankingRangesCons
        {
            get { return _riskRankingRangesCons; }
            set { _riskRankingRangesCons = value; }
        }

        public Configuration Configuration
        {
            get { return _configuration; }
            set { _configuration = value; }
        }

        // Configuration




        //------------------------------------------------------------------------------
        // Get the System name from the SystemID
        //------------------------------------------------------------------------------
        public string system_GetSystemName(int SystemID)
        {
            return GetNameFromID(SystemID.ToString(), "SystemID", "System", "Dbf Systems");
        }

        //------------------------------------------------------------------------------
        // Get the Line name from the LineID
        //------------------------------------------------------------------------------
        public string line_GetLineName(int LineID)
        {
            return GetNameFromID(LineID.ToString(), "LineID", "Line", "Dbf Lines");
        }

        //------------------------------------------------------------------------------
        // Get the Line Section name from the Zone_ID
        //------------------------------------------------------------------------------
        public string linesection_GetLineSectionName(int Zone_ID)
        {
            return GetNameFromID(Zone_ID.ToString(), "Zone ID", "Line Section", "Dbf Line Section");
        }


        //------------------------------------------------------------------------------
        // Get the Line Section names from the Zone_ID array
        //------------------------------------------------------------------------------
        public int linesection_GetLineSectionNameList(ref int[] ZoneIDList, ref string[] LSNames)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            int counter;
            if (connection != null)
            {
                connection.Open();
                for (counter = 0; counter <= (ZoneIDList.Length - 1); counter++)
                {
                    string queryString = "SELECT [Line Section],[Zone ID] from [Dbf Line Section] WHERE [Zone ID] = " + ZoneIDList[counter] ;
                    OleDbCommand command = new OleDbCommand(queryString, connection);
                    try
                    {
                        OleDbDataReader reader = command.ExecuteReader();
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                    }

                }
                connection.Close();
                return counter;
            }
            else
                return -1;
        }


        //------------------------------------------------------------------------------
        // Get Line Section Statistics
        //------------------------------------------------------------------------------
        public bool linesectionGetStatistics(ref int totalLength,ref List<string> unitlengths)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    // Get total length
                    string cmdStr = "SELECT Sum([Dbf Line Section].[Total Length]) AS [TotalLength] FROM [Dbf Line Section]";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    OleDbDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        totalLength = Convert.ToInt32(reader[0].ToString());
                    }
                    reader.Close();

                    // Get length by unit
                    List<int> units = new List<int>();
                    if (unitlengths == null) return false;

                    unitlengths.Clear();
                    int ulength = 0;
                    if (!GetDistinctSystemUnits(ref units))
                        return false;
                    int i;
                    for (i = 0; i <= units.Count - 1; i++)
                    {
                        cmdStr = "SELECT Sum([Dbf Line Section].[Total Length]) AS TotalLength, [Dbf Line Section].Unit FROM [Dbf Line Section] GROUP BY [Dbf Line Section].Unit HAVING ((([Dbf Line Section].Unit)='" + units[i].ToString() +"'))";
                        cmd = new OleDbCommand(cmdStr, connection);
                        reader = cmd.ExecuteReader();
                        if (reader.Read())
                        {
                            ulength = Convert.ToInt32(reader[0].ToString());
                            unitlengths.Add(string.Format("{0}:{1}", units[i].ToString(), ulength.ToString()));
                        }
                        reader.Close();
                    }

                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }



        //------------------------------------------------------------------------------
        // Returns a list of the units available
        //------------------------------------------------------------------------------
        public bool GetDistinctSystemUnits(ref List<int> Units)
        {
            Units.Clear();
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string queryString = "SELECT DISTINCT [Dbf Systems].Unit FROM [Dbf Systems]";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        Units.Add(Convert.ToInt32(reader[0].ToString()));
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    return false;
                }
                connection.Close();
                return (Units.Count > 0);
            }
            else
                return false;
        }


        //------------------------------------------------------------------------------
        // Load the FeatureDescription values in to the Select Set Item List
        //------------------------------------------------------------------------------
        public bool SetSelectionSetFeatureDescriptions(List<SelectedItemData> SelectedFeatures, LayerDefs Layers)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            int counter = 0;
            bool itemRemoved = false;

            if (connection != null)
            {
                connection.Open();
                // Loop through the items
                while (counter <= (SelectedFeatures.Count - 1))
                //for (counter = 0; counter <= (SelectedFeatures.Count - 1); counter++)
                {
                    SelectedItemData selItem = SelectedFeatures[counter];
                    // Build query to get the description
                    string queryString = "SELECT [" + Layers.CWFeatureDisplayField(selItem.layername) + "],[" + Layers.CWIndexName(selItem.layername) + "] from [" + Layers.CWTableName(selItem.layername) + "] WHERE [" + Layers.CWIndexName(selItem.layername) + "] = ";
                    if (Layers.KeyFieldIsString(selItem.layername))
                        queryString = queryString  + "\"" + selItem.CWKeyValue + "\"";
                    else
                        queryString = queryString + selItem.CWKeyValue;
                    OleDbCommand command = new OleDbCommand(queryString, connection);
                    try
                    {
                        OleDbDataReader reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            selItem.featuredescription = reader[0].ToString();
                        }
                        else
                        {
                            SelectedFeatures.Remove(selItem);
                            itemRemoved = true;
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                    }
                    if (!itemRemoved)
                        counter++;
                    itemRemoved = false;
                }
                connection.Close();
                return true;
            }
            else
                return false;
        }


        //------------------------------------------------------------------------------
        // Load the selected items list for the specified whereclause and layer
        //------------------------------------------------------------------------------
        public string LoadSelectedItemsList(string WhereClause, LayerDef Layer, bool ClearExisting, ref List<SelectedItemData> SelectedFeatures)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            string itemname = "";
            if (connection != null)
            {
                // Clear any selected items in the list if needed
                if (ClearExisting)
                    SelectedFeatures.Clear();
                SelectedItemData selitem;
                // Run the query
                connection.Open();
                string queryString = string.Format("SELECT [{0}], [{1}], MapFeatureID FROM [{2}] WHERE {3}",Layer.GIS_CW_Index_Field, Layer.GIS_CW_Feature_Display_Field, Layer.GIS_CW_Table_Name, WhereClause);
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    OleDbDataReader reader = command.ExecuteReader();                    
                    while (reader.Read())
                    {
                        selitem = new SelectedItemData();
                        selitem.AttribObjectID = Convert.ToInt32(reader[2].ToString());
                        selitem.CWKeyValue = reader[0].ToString();
                        selitem.featuredescription = reader[1].ToString();
                        selitem.layername = Layer.layername;
                        selitem.nodelevel = 1;
                        SelectedFeatures.Add(selitem);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                    return "";
                }
                connection.Close();
                return itemname;
            }
            else
                return "";
        }


        #region Update Feature Cross References

        //------------------------------------------------------------------------------
        // Set Line Record Cross Reference
        //------------------------------------------------------------------------------
        public bool lineSetRecordCrossReference(int CWID, int ObjectID)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = string.Format("UPDATE [Dbf Lines] SET [Dbf Lines].MapFeatureID = {0} WHERE ((([Dbf Lines].LineID)={1}))", ObjectID, CWID);
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }

        //------------------------------------------------------------------------------
        // Clear Line Record Cross References
        //------------------------------------------------------------------------------
        public bool lineClearRecordCrossReferences()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = "UPDATE [Dbf Lines] SET [Dbf Lines].MapFeatureID = Null;";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }



        //------------------------------------------------------------------------------
        // Clear Line Section Record Cross References
        //------------------------------------------------------------------------------
        public bool linesectionClearRecordCrossReferences()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = "UPDATE [Dbf Line Section] SET [Dbf Line Section].MapFeatureID = Null;";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }



        //------------------------------------------------------------------------------
        // Set Line Section Record Cross Reference
        //------------------------------------------------------------------------------
        public bool linesectionSetRecordCrossReference(int CWID, int ObjectID)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = string.Format("UPDATE [Dbf Line Section] SET [Dbf Line Section].MapFeatureID = {0} WHERE ((([Dbf Line Section].[Zone ID])={1}))", ObjectID, CWID);
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }


        //------------------------------------------------------------------------------
        // Set Boring Location Record Cross Reference
        //------------------------------------------------------------------------------
        public bool blSetRecordCrossReference(string CWID, int ObjectID)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = string.Format("UPDATE Soil_Boring_Locations SET Soil_Boring_Locations.MapFeatureID = {0} WHERE (((Soil_Boring_Locations.Boring_Location_ID)=\"{1}\"))", ObjectID, CWID);
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }


        //------------------------------------------------------------------------------
        // Clear Boring Location Record Cross References
        //------------------------------------------------------------------------------
        public bool blClearRecordCrossReferences()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = "UPDATE Soil_Boring_Locations SET Soil_Boring_Locations.MapFeatureID = Null;";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }


        //------------------------------------------------------------------------------
        // Set Corrosion Probe Record Cross Reference
        //------------------------------------------------------------------------------
        public bool cpSetRecordCrossReference(string CWID, int ObjectID)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = string.Format("UPDATE Soil_Corrosion_Probes SET Soil_Corrosion_Probes.MapFeatureID = {0} WHERE (((Soil_Corrosion_Probes.Corrosion_Probe_ID)=\"{1}\"))", ObjectID, CWID);
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }


        //------------------------------------------------------------------------------
        // Clear Corrosion Probe Record Cross References
        //------------------------------------------------------------------------------
        public bool cpClearRecordCrossReferences()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = "UPDATE Soil_Corrosion_Probes SET Soil_Corrosion_Probes.MapFeatureID = Null;";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }



        //------------------------------------------------------------------------------
        // Set Rectifier Record Cross Reference
        //------------------------------------------------------------------------------
        public bool rectSetRecordCrossReference(string CWID, int ObjectID)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = string.Format("UPDATE [Dbf Rectifiers] SET [Dbf Rectifiers].[MapFeatureID] = {0} WHERE ((([Dbf Rectifiers].RectifierID)={1}))", ObjectID, CWID);
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }


        //------------------------------------------------------------------------------
        // Clear Rectifier Record Cross References
        //------------------------------------------------------------------------------
        public bool rectClearRecordCrossReferences()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = "UPDATE [Dbf Rectifiers] set [Dbf Rectifiers].[MapFeatureID] = Null;";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }


        //------------------------------------------------------------------------------
        // Clear Test Station Record Cross References
        //------------------------------------------------------------------------------
        public bool teststationClearRecordCrossReferences()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = "UPDATE [Dbf Stations] set [Dbf Stations].[MapFeatureID] = Null;";
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }


        #endregion


        //------------------------------------------------------------------------------
        // SendData To CompWorks - Associate the Line Selected Line Section to the Boring Location
        // FeatureKey: 1001
        //------------------------------------------------------------------------------
        public string linesection_SetAssociationToBoringLocation(int Zone_ID, string BLID)
        {
            string itemname = "";
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string cmdStr = "Select count(*) from [Soil - Line Section] where [Zone ID] = " + Zone_ID; //get the existence of the record as count 
                OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                int count = (int)cmd.ExecuteScalar();
                if (count == 0)
                {
                    string queryString = "INSERT INTO [Soil - Line Section] ([Zone ID], Boring_Location_ID ) Values (" + Zone_ID + " , '" + BLID + "');";
                    OleDbCommand command = new OleDbCommand(queryString, connection);
                    try
                    {
                        OleDbDataReader reader = command.ExecuteReader();
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        return "";
                    }
                    //record already exist 
                }
                return itemname;
            }
            return itemname;
        }


         //------------------------------------------------------------------------------
        // Get Item Name from ItemID
        // Parameters: 
        //  > ItemID - the table Key/ID field to search
        //  > ValueField - the data field value to return
        //  > TableName - the table to search
        //------------------------------------------------------------------------------
        public string GetNameFromID(string ItemID, string KeyFieldName, string ValueFieldName, string TableName)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            string itemname = "";
            if (connection != null)
            {
                connection.Open();
                string queryString = "SELECT [" + ValueFieldName + "] from [" + TableName + "] WHERE [" + KeyFieldName + "] = " + ItemID.ToString();
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        itemname = reader[0].ToString();
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                    return "";
                }
                connection.Close();
                return itemname;
            }
            else
                return "";
        }



        //------------------------------------------------------------------------------
        // Load the GIS Settings
        //------------------------------------------------------------------------------
        public bool setup_Get_GIS_Settings(ref GISDBSettings GIS_Settings)
        {
            string strval = "";
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string queryString = "SELECT GIS_Settings.Settings_ID, GIS_Settings.GIS_Enabled, GIS_Settings.GIS_Max_Easting, GIS_Settings.GIS_Min_Easting, GIS_Settings.GIS_Max_Northing, GIS_Settings.GIS_Min_Northing, GIS_Settings.DBRootFolder, GIS_Settings.SampleSelectionByGroundwater, GIS_Settings.GIS_MX_File_Location, GIS_Settings.GIS_Exe_Path, GIS_Settings.CWGraph_Exe_Path, GIS_Settings.CW_System_File_Path FROM GIS_Settings";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    OleDbDataReader reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        GIS_Settings.GISRootFolder = reader[6].ToString();
                        strval = reader[3].ToString();
                        if (strval.Length>0) GIS_Settings.MinEasting = Convert.ToDouble(strval);
                        strval = reader[2].ToString();
                        if (strval.Length > 0) GIS_Settings.MaxEasting = Convert.ToDouble(strval);
                        strval = reader[5].ToString();
                        if (strval.Length > 0) GIS_Settings.MinNorthing = Convert.ToDouble(strval);
                        strval = reader[4].ToString();
                        if (strval.Length > 0) GIS_Settings.MaxNorthing = Convert.ToDouble(strval);
                        strval = reader[7].ToString();
                        if (strval.Length > 0) GIS_Settings.SampleSelectionByGroundwater = Convert.ToBoolean(strval);
                        strval = reader[8].ToString();
                        if (strval.Length > 0) GIS_Settings.GISmxdfile = strval;
                        strval = reader[10].ToString();
                        if (strval.Length > 0) GIS_Settings.CWGraphingEXE = strval;
                        reader.Close();                        
                    }
                }
                catch (Exception ex)
                {
                    connection.Close();
                    return false;
                }
                connection.Close();
                return true;
            } // if connection <> null
            else
                return false;
        }


        //------------------------------------------------------------------------------
        // Get Data From CompWorks - Associate the Line and System
        //------------------------------------------------------------------------------
        public void line_GetLinesInSystem(string system, List<string> LineIDs)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string cmdStr = "Select LineID from [Dbf System_Line] where [SystemID] = " + system; //get the existence of the record as count 
                OleDbCommand command = new OleDbCommand(cmdStr, connection);
                int count = (int)command.ExecuteScalar();
                if (count > 0)
                {
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        LineIDs.Add(reader[0].ToString());
                    }
                    reader.Close();
                }
                connection.Close();
                }
            }


        //------------------------------------------------------------------------------
        // Add a Filter
        //------------------------------------------------------------------------------
        public bool filter_AddNew(SQLFilterClass filter)
        {
            if (filter == null) return false;
            if (filter.ID > 0) return false;

            int FilterID = 0;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                // First add the GIS_Filters record
                string queryString = "INSERT INTO [GIS_Filters] (GIS_Filter_Name, GIS_Filter_Description,  GIS_Filter_Layer_Index) Values ('" + filter.FilterName + "' , '" + filter.FilterDescription + "' , " + filter.LayerIndex + ");";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    OleDbDataReader reader = command.ExecuteReader();
                    reader.Close();
                }
                catch (Exception ex)
                {
                    return false;
                }

                // Get the ID of the filter just added so we can set the FK for the conditions records
                command.CommandText = "Select @@Identity";
                string strquote = "";
                FilterID = (int)command.ExecuteScalar();
                if (FilterID > 0)
                {
                    filter.ID = FilterID;
                    // Next add the GIS_Filter_Cond record(s)
                    int i;
                    for (i = 0; i <= (filter.ConditionCount - 1); i++)
                    {
                        SQLCondition cond = new SQLCondition();
                        if (filter.GetCondition(i, ref cond))
                        {
                            // if a character field, add quotes to the value
                            if (cond.fieldtype == "DBTYPE_WVARCHAR")
                                strquote = "'";
                            else
                                strquote = "";
                            if (cond.logicaloperator.Length != 0)
                                queryString = "INSERT INTO [GIS_Filter_Cond] (GIS_Filters_ID, GIS_Filter_Cond_FieldName, GIS_Filter_Cond_Condition, GIS_Filter_Cond_Value, GIS_Filter_Cond_FieldType, GIS_Filter_Cond_Operator) " +
                                    "Values (" + FilterID + " , '" + cond.fieldname + "' , '" + cond.condition + "' , " + strquote + cond.value + strquote + " , '" + cond.fieldtype + "', '" + cond.logicaloperator + "')";
                            else
                                queryString = "INSERT INTO [GIS_Filter_Cond] (GIS_Filters_ID, GIS_Filter_Cond_FieldName, GIS_Filter_Cond_Condition, GIS_Filter_Cond_Value, GIS_Filter_Cond_FieldType) " +
                                    "Values (" + FilterID + " , '" + cond.fieldname + "' , '" + cond.condition + "' , " + strquote + cond.value + strquote + " , '" + cond.fieldtype + "')";
                            command = new OleDbCommand(queryString, connection);
                            try
                            {
                                OleDbDataReader reader = command.ExecuteReader();
                                reader.Close();
                            }
                            catch (Exception ex)
                            {
                                return false;
                            }
                        }
                    }
                }
            }
            return true;
        }


        //------------------------------------------------------------------------------
        // Replace a Filter
        //------------------------------------------------------------------------------
        public bool filter_Update(SQLFilterClass filter)
        {
            if (filter == null) return false;
            if (filter.ID <= 0) return false;

            // It's assumed that it's an existing filter with an established ID
            // Need to find the ID and remove existing filter and condition records
            // then need to save the new one
            int FilterID = filter.ID;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                OleDbDataReader reader;

                // DELETE ORIGINAL FILTER
                // First remove the existing GIS_Filters record
                string queryString =
                    "DELETE GIS_Filters.GIS_Filters_ID FROM GIS_Filters WHERE (((GIS_Filters.GIS_Filters_ID)=" +
                    FilterID.ToString() + "))";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    command.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    return false;
                }

                // Next remove the existing GIS_Filter_Conditions record
                queryString = "DELETE GIS_Filter_Cond.GIS_Filters_ID FROM GIS_Filter_Cond WHERE (((GIS_Filter_Cond.GIS_Filters_ID)=" +
                    FilterID.ToString() + "))";
                command = new OleDbCommand(queryString, connection);
                try
                {
                    command.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    return false;
                }

                // SAVE NEW FILTER
                // Now save the new filter
                queryString = "INSERT INTO [GIS_Filters] (GIS_Filter_Name, GIS_Filter_Description,  GIS_Filter_Layer_Index) Values ('" + filter.FilterName + "' , '" + filter.FilterDescription + "' , " + filter.LayerIndex + ");";
                command = new OleDbCommand(queryString, connection);
                try
                {
                    reader = command.ExecuteReader();
                    reader.Close();
                }
                catch (Exception ex)
                {
                    return false;
                }

                // Get the ID of the filter just added so we can set the FK for the conditions records
                command.CommandText = "Select @@Identity";
                string strquote = "";
                FilterID = (int)command.ExecuteScalar();
                if (FilterID > 0)
                {
                    filter.ID = FilterID;
                    // Next add the GIS_Filter_Cond record(s)
                    int i;
                    for (i = 0; i <= (filter.ConditionCount - 1); i++)
                    {
                        SQLCondition cond = new SQLCondition();
                        if (filter.GetCondition(i, ref cond))
                        {
                            // if a character field, add quotes to the value
                            if (cond.fieldtype == "DBTYPE_WVARCHAR")
                                strquote = "'";
                            else
                                strquote = "";
                            if (cond.logicaloperator.Length != 0)
                                queryString = "INSERT INTO [GIS_Filter_Cond] (GIS_Filters_ID, GIS_Filter_Cond_FieldName, GIS_Filter_Cond_Condition, GIS_Filter_Cond_Value, GIS_Filter_Cond_FieldType, GIS_Filter_Cond_Operator) " +
                                    "Values (" + FilterID + " , '" + cond.fieldname + "' , '" + cond.condition + "' , " + strquote + cond.value + strquote + " , '" + cond.fieldtype + "', '" + cond.logicaloperator + "')";
                            else
                                queryString = "INSERT INTO [GIS_Filter_Cond] (GIS_Filters_ID, GIS_Filter_Cond_FieldName, GIS_Filter_Cond_Condition, GIS_Filter_Cond_Value, GIS_Filter_Cond_FieldType) " +
                                    "Values (" + FilterID + " , '" + cond.fieldname + "' , '" + cond.condition + "' , " + strquote + cond.value + strquote + " , '" + cond.fieldtype + "')";
                            command = new OleDbCommand(queryString, connection);
                            try
                            {
                                reader = command.ExecuteReader();
                                reader.Close();
                            }
                            catch (Exception ex)
                            {
                                return false;
                            }
                        }
                    }
                }
            }
            return true;
        }


        //------------------------------------------------------------------------------
        // Remove a Filter
        //------------------------------------------------------------------------------
        public bool filter_Remove(int FilterID)
        {
            if (FilterID <= 0) return false;

            // It's assumed that it's an existing filter with an established ID
            // Need to find the ID and remove existing filter and condition records
            // then need to save the new one
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                OleDbDataReader reader;

                // DELETE ORIGINAL FILTER
                // First remove the existing GIS_Filters record
                string queryString =
                    "DELETE GIS_Filters.GIS_Filters_ID FROM GIS_Filters WHERE (((GIS_Filters.GIS_Filters_ID)=" +
                    FilterID.ToString() + "))";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    command.ExecuteScalar();                    
                }
                catch (Exception ex)
                {
                    return false;
                }

                // Next remove the existing GIS_Filter_Conditions record
                queryString = "DELETE GIS_Filter_Cond.GIS_Filters_ID FROM GIS_Filter_Cond WHERE (((GIS_Filter_Cond.GIS_Filters_ID)=" +
                    FilterID.ToString() + "))";
                command = new OleDbCommand(queryString, connection);
                try
                {
                    command.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            return true;
        }


        //------------------------------------------------------------------------------
        // Load the Layer Saved Filters
        //------------------------------------------------------------------------------
        public bool LoadLayerSavedFilters(LayerDef layer)
        {
            if (layer == null) return false;
            layer.SavedFilters.Clear();

            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string queryString = "SELECT  GIS_Filters.GIS_Filters_ID, GIS_Filters.GIS_Filter_Name, GIS_Filters.GIS_Filter_Description, GIS_Filters.GIS_Filter_Layer_Index " +
                    "FROM GIS_Filters WHERE (((GIS_Filters.GIS_Filter_Layer_Index)=" + layer.SystemUse + "));";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        // first get the layer information
                        SQLFilterClass newfilter = new SQLFilterClass();
                        newfilter.ID = Convert.ToInt32(reader[0].ToString());
                        newfilter.FilterName = reader[1].ToString();
                        newfilter.FilterDescription = reader[2].ToString();
                        newfilter.LayerIndex = (LayerSystemUse)Convert.ToInt32(reader[3].ToString());

                        // Now get the condition records
                        OleDbConnection connection2 = new OleDbConnection(_connectionstring);
                        if (connection2 != null)
                        {
                            connection2.Open();
                            string queryCond = "SELECT " + 
                                                "GIS_Filter_Cond.GIS_Filter_Cond_ID, " +            //0
                                                "GIS_Filter_Cond.GIS_Filters_ID, " +                //1
                                                "GIS_Filter_Cond.GIS_Filter_Cond_Operator, " +      //2
                                                "GIS_Filter_Cond.GIS_Filter_Cond_FieldName, " +     //3
                                                "GIS_Filter_Cond.GIS_Filter_Cond_Condition, " +     //4
                                                "GIS_Filter_Cond.GIS_Filter_Cond_Value, " +         //5
                                                "GIS_Filter_Cond.GIS_Filter_Cond_FieldType " +      //6
                                                "FROM GIS_Filter_Cond " +
                                                "WHERE (((GIS_Filter_Cond.GIS_Filters_ID)=" + newfilter.ID.ToString() + "));";
                            if (newfilter != null)
                            {
                                OleDbCommand command2 = new OleDbCommand(queryCond, connection2);
                                try
                                {
                                    OleDbDataReader reader2 = command2.ExecuteReader();
                                    while (reader2.Read())
                                    {
                                        SQLCondition cond = new SQLCondition();
                                        cond.logicaloperator = reader2[2].ToString();
                                        cond.fieldname = reader2[3].ToString();
                                        cond.condition = reader2[4].ToString();
                                        cond.value = reader2[5].ToString();
                                        cond.fieldtype = reader2[6].ToString();
                                        newfilter.AddCondition(cond);
                                    }
                                    reader2.Close();
                                }
                                catch (Exception ex)
                                {
                                    connection2.Close();
                                    return false;
                                }
                            } // newfilter not null

                        } // Condition connection not null
                        layer.AddSavedFilter(newfilter);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                    return false;
                }
                connection.Close();
                return true;
            }
            else
                return false;
        }
      


        //------------------------------------------------------------------------------
        // Load the Layer Definitions
        //------------------------------------------------------------------------------
        public bool SetLayerDefs(LayerDefs _layers)
        {
            int count = 0;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string queryString = "SELECT GIS_Layers.GIS_Layer_Name, GIS_Layers.GIS_Layer_Description, GIS_Layers.GIS_Layer_Visible_At_Startup, GIS_Layers.GIS_Layer_SystemUse, "+
                "GIS_Layers.GIS_CW_Table_Name, GIS_Layers.GIS_CW_Feature_Display_Field, GIS_Layers.GIS_CW_Index_Field, GIS_Layers.GIS_Key_Field_Is_String, GIS_Layers.GIS_Attribute_Index_Field, GIS_Layers.GIS_Layers_ID FROM GIS_Layers";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    LayerDef ld;
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        ld = new LayerDef();
                        ld.layername = GetFieldDataStr("GIS_Layer_Name", ref reader);
                        ld.Description = GetFieldDataStr("GIS_Layer_Description", ref reader);
                        ld.LayerVisibleAtStartup= Convert.ToBoolean(GetFieldDataStr("GIS_Layer_Visible_At_Startup", ref reader));
                        ld.SystemUse = (LayerSystemUse)Convert.ToInt32(GetFieldDataInt("GIS_Layer_SystemUse", ref reader));
                        ld.GIS_CW_Table_Name = GetFieldDataStr("GIS_CW_Table_Name", ref reader);
                        ld.GIS_CW_Feature_Display_Field = GetFieldDataStr("GIS_CW_Feature_Display_Field", ref reader);
                        ld.GIS_CW_Index_Field = GetFieldDataStr("GIS_CW_Index_Field", ref reader);
                        ld.GIS_Key_Field_Is_String = Convert.ToBoolean(GetFieldDataStr("GIS_Key_Field_Is_String", ref reader));
                        ld.GIS_Attribute_Index_Field = GetFieldDataStr("GIS_Attribute_Index_Field", ref reader);
                        ld.GISLayerID = GetFieldDataInt("GIS_Layers_ID", ref reader);

                        //ld.layername = reader[0].ToString();
                        //ld.Description = reader[1].ToString();
                        //ld.LayerVisibleAtStartup = Convert.ToBoolean(reader[2].ToString());
                        //ld.SystemUse = (LayerSystemUse)Convert.ToInt32(reader[3].ToString());
                        //ld.GIS_CW_Table_Name = reader[4].ToString();
                        //ld.GIS_CW_Feature_Display_Field = reader[5].ToString();
                        //ld.GIS_CW_Index_Field = reader[6].ToString();
                        //ld.GIS_Key_Field_Is_String = Convert.ToBoolean(reader[7].ToString());
                        //ld.GIS_Attribute_Index_Field = reader[8].ToString();
                        //ld.GISLayerID = Convert.ToInt32(reader[9].ToString());
                       
                        LoadLayerSavedFilters(ld);
                        _layers.AddLayer(ld);
                        count++;
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                    return false;
                }
                connection.Close();
                return true;
            }
            else
                return false;
        }


        //------------------------------------------------------------------------------
        // Execute a filter and return the key values in the List<string> results
        //------------------------------------------------------------------------------
        public bool ExecuteFilter(SQLFilterClass filter,string basequery, string keyfieldname, ref List<string> results)
        {
            if (filter == null) return false;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                // Set the sql command to the base query
                string cmdStr = basequery;
                // Add the conditions (WHERE)
                cmdStr += " WHERE (";
                int i;
                SQLCondition cond = new SQLCondition();
                for (i = 0; i <= (filter.ConditionCount - 1); i++)
                {
                    if (filter.GetCondition(i, ref cond))
                    {
                        if ((i>0) && (cond.logicaloperator.Length > 0)) 
                            cmdStr += string.Format(" {0} ", cond.logicaloperator);
                        if (cond.fieldtype == "DBTYPE_WVARCHAR")
                            cmdStr += string.Format("([{0}] {1} '{2}')", cond.fieldname, cond.condition, cond.value);
                        else
                            cmdStr += string.Format("([{0}] {1} {2})", cond.fieldname, cond.condition, cond.value);
                    }
                }
                cmdStr += ")";
                int mycount = 0;
                int keyidx = 0;;

                OleDbDataReader reader;
                OleDbCommand command = new OleDbCommand(cmdStr, connection);
                try
                {
                    reader = command.ExecuteReader();
                    keyidx = reader.GetOrdinal(keyfieldname);
                    while (reader.Read())
                    {
                        if (reader.FieldCount > 0)
                        {
                            results.Add(reader[keyidx].ToString());
                            mycount++;
                        }
                    }
                    connection.Close();
                    return true;
                }
                catch (Exception ex)
                {
                }
                connection.Close();
            }
            return false;
        }


        //------------------------------------------------------------------------------
        // Execute SQL Statement and populate the selected items list
        //------------------------------------------------------------------------------
        public bool ExecuteSQLtoSelectionSet(string SQLStatement, LayerDef ldef, ref List<SelectedItemData> SelectionSet)
        {
            if (SQLStatement.Length == 0) return false;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                // Set the sql command to the base query
                string cmdStr = SQLStatement;
                int keyidx = 0; ;
                int descidx = 0;
                int attkeyidx = 0;
                SelectedItemData selItem; ;
                OleDbDataReader reader;
                OleDbCommand command = new OleDbCommand(cmdStr, connection);
                try
                {
                    reader = command.ExecuteReader();
                    keyidx = reader.GetOrdinal(ldef.GIS_CW_Index_Field);
                    descidx = reader.GetOrdinal(ldef.GIS_CW_Feature_Display_Field);
                    attkeyidx = reader.GetOrdinal(CW_XRef_Fieldname);
                    while (reader.Read())
                    {
                        if (reader[attkeyidx].ToString().Length > 0)
                        {
                            selItem = new SelectedItemData();
                            selItem.CWKeyValue = reader[keyidx].ToString();
                            selItem.AttribObjectID = Convert.ToInt32(reader[attkeyidx].ToString());
                            selItem.featuredescription = reader[descidx].ToString();
                            selItem.nodelevel = 1;
                            selItem.layername = ldef.layername;
                            SelectionSet.Add(selItem);                            
                        }
                    }
                    connection.Close();
                    return true;
                }
                catch (Exception ex)
                {
                }
                connection.Close();
            }
            return false;
        }


        //------------------------------------------------------------------------------
        // Get the Index for a specified Generic Type
        //------------------------------------------------------------------------------
        public int getGenericTypeID(string GenericType)
        {
            int genericid = 0;
            string SQLStatement =
                "SELECT [Dbf Setup -  Data Generic].GenericID, [Dbf Setup -  Data Generic].GenericType FROM [Dbf Setup -  Data Generic] WHERE ((([Dbf Setup -  Data Generic].GenericType)=\"" +
                GenericType + "\"));";
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                // Set the sql command to the base query
                string cmdStr = SQLStatement;
                OleDbDataReader reader;
                OleDbCommand command = new OleDbCommand(cmdStr, connection);
                try
                {
                    reader = command.ExecuteReader();
                    int keyidx = reader.GetOrdinal("GenericID");
                    if (reader.Read())
                    {
                        genericid = (int)reader[keyidx];
                    }
                    connection.Close();
                }
                catch (Exception ex)
                {

                }
                connection.Close();
            }
            return genericid;
        }

        //------------------------------------------------------------------------------
        // Get the Index for a specified Setup Data Type
        //------------------------------------------------------------------------------
        public int getSetupDataTypeID(string SetupDataType, string ValueData)
        {
            int setupid = 0;
            string SQLStatement =
                "SELECT [Dbf Setup -  Data].SetupID, [Dbf Setup -  Data].Type, [Dbf Setup -  Data].ValueData FROM [Dbf Setup -  Data] WHERE ([Dbf Setup -  Data].ValueData =\"" +
                ValueData + "\" AND [Dbf Setup -  Data].Type=\"" + SetupDataType + "\")";
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                // Set the sql command to the base query
                string cmdStr = SQLStatement;
                OleDbDataReader reader;
                OleDbCommand command = new OleDbCommand(cmdStr, connection);
                try
                {
                    reader = command.ExecuteReader();
                    int keyidx = reader.GetOrdinal("SetupID");
                    if (reader.Read())
                    {
                        setupid = (int)reader[keyidx];
                    }
                    connection.Close();
                }
                catch (Exception ex)
                {

                }
                connection.Close();
            }
            return setupid;
        }

                    
        //------------------------------------------------------------------------------
        // Find Line Section Inspections
        //------------------------------------------------------------------------------
        public bool find_Line_Section_Inspections(ref List<SelectedItemData> SelectionSet, LayerDef ldef, int InspectionSetupID)
        {
            // Build the query to search with
            string SQLStatement = "SELECT [Dbf Line Section].[Zone ID], [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point], [Dbf Line Section].[Line Section], [Dbf Line Section].[MapFeatureID], " +
                                  "[Dbf Inspection Range].Method FROM [Dbf Inspection Range] " +
                                  "INNER JOIN [Dbf Line Section] ON [Dbf Inspection Range].LineID = [Dbf Line Section].LineID " +
                                  "WHERE ((([Dbf Line Section].[Begin Point])>=[Dbf Inspection Range].[Begin Point]) AND (([Dbf Line Section].[End Point])<=[Dbf Inspection Range].[End Point]) AND (([Dbf Inspection Range].Method)=" + InspectionSetupID + "))";
            return ExecuteSQLtoSelectionSet(SQLStatement, ldef, ref SelectionSet);
        }


        //------------------------------------------------------------------------------
        // Find Line Section Evaluations
        //------------------------------------------------------------------------------
        public bool find_Line_Section_Evaluations(ref List<SelectedItemData> SelectionSet, LayerDef ldef, string Method, string WorkScope)
        {
            // Build the query to search with
            string SQLStatement =
                "SELECT [Dbf Inspections].[Zone ID], [Dbf Inspections].Method, [Dbf Inspections].[WS Name], [Dbf Line Section].[Line Section], [Dbf Line Section].MapFeatureID " +
                "FROM [Dbf Inspections] INNER JOIN [Dbf Line Section] ON [Dbf Inspections].[Zone ID] = [Dbf Line Section].[Zone ID]";

            string whereclause = "";
            if (Method.Length > 0)
            {
                if (string.Compare(Method, "All") != 0)
                {
                    whereclause = "([Dbf Inspections].Method ='" + Method + "')";
                }
            }
            if (WorkScope.Length > 0)
            {
                if (string.Compare(WorkScope, "All") != 0)
                {
                    if (whereclause.Length > 0)
                        whereclause += " OR ";
                    whereclause = "([Dbf Inspections].[WS Name] = '" + WorkScope + "')";
                }
            }
            if (whereclause.Length > 0)
            {
                SQLStatement += " WHERE " + whereclause;
            }            
            return ExecuteSQLtoSelectionSet(SQLStatement, ldef, ref SelectionSet);
        }


        //------------------------------------------------------------------------------
        // Find Line Sections with LCO
        // LCO = (0=None, 5,4,3,2,1 = 30, 14, 7 day etc...
        // LCO = <0 means show any with LCO at all
        //------------------------------------------------------------------------------
        public bool find_Line_Sections_with_LCO(ref List<SelectedItemData> SelectionSet, LayerDef ldef, int LCO, bool inclusivebelow)
        {
            // Build the query to search with
            string LCOstr = "";
            if (LCO < 0) // Do any with LCO
                LCOstr = "(([Dbf Consequence].LCO) > 0)";
            else
                if ((inclusivebelow) && (LCO > 1))
                {
                    LCOstr = "((([Dbf Consequence].LCO) <= " + LCO.ToString() + ") AND (([Dbf Consequence].LCO) > 0))";
                }
                else
                {
                    LCOstr = "(([Dbf Consequence].LCO) = " + LCO.ToString() + ")";
                }
            string SQLStatement =
                "SELECT [Dbf Line Section].[Zone ID], [Dbf Consequence].LCO, [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point], " +
                "[Dbf Line Section].[Line Section], [Dbf Line Section].MapFeatureID " +
                "FROM [Dbf Consequence] INNER JOIN [Dbf Line Section] ON [Dbf Consequence].LineID = [Dbf Line Section].LineID " +
                "WHERE (" + LCOstr + " AND (([Dbf Line Section].[Begin Point])>=[Dbf Consequence].[Begin Point]) AND (([Dbf Line Section].[End Point])<=[Dbf Consequence].[End Point]))";
            return ExecuteSQLtoSelectionSet(SQLStatement, ldef, ref SelectionSet);
        }



        //------------------------------------------------------------------------------
        // Returns true if a specified InspectionID has GWUT results at the specified level  
        //------------------------------------------------------------------------------
        public bool find_GWUT_Line_Section_has_Level(int InspectionID, int SelLevel)
        {
            bool retval = false;
            string SQLStatement = "SELECT [Dbf Inspections Guided Wave].InspectionID, [Dbf Inspections Guided Wave].ResultsLevel, " +
                                  "[Dbf Inspections Guided Wave].ResultClass FROM [Dbf Inspections Guided Wave] " +
                                  "WHERE ((([Dbf Inspections Guided Wave].ResultsLevel)=" + SelLevel.ToString() + "))";
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            connection.Open();
            // Set the sql command to the base query
            string cmdStr = SQLStatement;
            OleDbDataReader reader;
            OleDbCommand command = new OleDbCommand(cmdStr, connection);
            try
            {
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    retval = true;
                }
                connection.Close();
            }
            catch (Exception ex)
            {
            }
            connection.Close();
            return retval;
        }


                //------------------------------------------------------------------------------
        // Find Line Sections by Risk Ranking
        //------------------------------------------------------------------------------
        public int line_section_Add_To_Selection(ref List<SelectedItemData> SelectionSet, LayerDef ldef, string SQLStatement)
        {
            int retval = 0;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            connection.Open();
            // Set the sql command to the base query
            string cmdStr = SQLStatement;
            OleDbDataReader reader;
            OleDbCommand command = new OleDbCommand(cmdStr, connection);
            try
            {
                SelectedItemData selItem;
                reader = command.ExecuteReader();
                int keyidx;
                string CWKeyValue;
                int attkeyidx = reader.GetOrdinal(CW_XRef_Fieldname);
                int descidx = reader.GetOrdinal(ldef.GIS_CW_Feature_Display_Field);
                int lstidx = 0;
                int numfound = 0;
                while (reader.Read())
                {
                    keyidx = reader.GetOrdinal(ldef.GIS_CW_Index_Field);
                    //                    keyidx = reader.GetOrdinal("Zone ID");
                    CWKeyValue = reader[keyidx].ToString();
                    // Check for duplicates
                    numfound = 0;
                    for (lstidx = 0; lstidx <= (SelectionSet.Count - 1); lstidx++)
                    {
                        if (SelectionSet[lstidx].CWKeyValue.ToString() == CWKeyValue) numfound++;
                    }
                    if (numfound == 0)
                    {

                        string strcheck = reader[attkeyidx].ToString();
                        if (strcheck.Length > 0)
                        {
                            selItem = new SelectedItemData();
                            selItem.CWKeyValue = CWKeyValue;
                            selItem.AttribObjectID = Convert.ToInt32(strcheck);
                            selItem.featuredescription = reader[descidx].ToString();
                            selItem.nodelevel = 1;
                            selItem.layername = ldef.layername;
                            SelectionSet.Add(selItem);
                        }
                    }
                }
                retval = SelectionSet.Count;
                connection.Close();
            }
            catch (Exception ex)
            {
                retval = 0;
            }
            return retval;            
        }


        //------------------------------------------------------------------------------
        // Find Line Sections by Risk Ranking
        //------------------------------------------------------------------------------
        public int find_Line_Section_by_Risk_Ranking(ref List<SelectedItemData> SelectionSet, LayerDef ldef, string sussql, string consql)
        {
            int retval = 0;
            if (sussql.Length > 0)
                retval = line_section_Add_To_Selection(ref SelectionSet, ldef, sussql);
            if (consql.Length > 0)
                retval = retval + line_section_Add_To_Selection(ref SelectionSet, ldef, consql);
            return retval;
        }


        //------------------------------------------------------------------------------
        // Find Line Sections with GWUT By Level
        //------------------------------------------------------------------------------
        public int find_Line_Section_GWUT_By_Level(ref List<SelectedItemData> SelectionSet, LayerDef ldef, int SelLevel)
        {
            int retval = 0;
            if (SelLevel <= 0) return retval;
            // Get the Guided Wave Inspection Set ID
            int InspectionSetupID = getSetupDataTypeID("InspectionMethod","GWUT");
            if (InspectionSetupID == 0) return 0;  // GWUT not found
            string SQLStatement =
                "SELECT [Dbf Line Section].[Zone ID], [Dbf Line Section].[Line Section], [Dbf Line Section].MapFeatureID, [Dbf Inspection Range].InspectionID, [Dbf Inspection Range].Method, [Dbf Inspections Guided Wave].ResultsLevel, [Dbf Inspections Guided Wave].[Begin Point], [Dbf Inspections Guided Wave].[End Point], [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point] FROM ([Dbf Inspection Range] INNER JOIN [Dbf Line Section] ON [Dbf Inspection Range].LineID = [Dbf Line Section].LineID) INNER JOIN [Dbf Inspections Guided Wave] ON [Dbf Inspection Range].InspectionID = [Dbf Inspections Guided Wave].InspectionID " +
                "WHERE ((([Dbf Inspection Range].Method)=67) AND (([Dbf Inspections Guided Wave].ResultsLevel)=\"" + SelLevel + "\") AND (([Dbf Line Section].[Begin Point])<=[Dbf Inspections Guided Wave].[Begin Point]) AND (([Dbf Line Section].[End Point])>=[Dbf Inspections Guided Wave].[End Point]))";

            // >=[Dbf Inspection Range].[Begin Point]

            OleDbConnection connection = new OleDbConnection(_connectionstring);
            connection.Open();
            // Set the sql command to the base query
            string cmdStr = SQLStatement;
            OleDbDataReader reader;
            OleDbCommand command = new OleDbCommand(cmdStr, connection);
            SelectionSet.Clear();
            try
            {
                SelectedItemData selItem;
                reader = command.ExecuteReader();
                int keyidx;
                string CWKeyValue;
                int attkeyidx = reader.GetOrdinal(CW_XRef_Fieldname);
                int descidx = reader.GetOrdinal(ldef.GIS_CW_Feature_Display_Field);
                int lstidx = 0;
                int numfound = 0;
                while (reader.Read())
                {
                    keyidx = reader.GetOrdinal(ldef.GIS_CW_Index_Field);
                    CWKeyValue = reader[keyidx].ToString();
                    // Check for duplicates
                    numfound = 0;
                    for (lstidx = 0; lstidx <= (SelectionSet.Count - 1); lstidx++)
                    {
                        if (SelectionSet[lstidx].CWKeyValue.ToString() == CWKeyValue) numfound++;
                    }
                    if (numfound == 0)
                    {

                        string strcheck = reader[attkeyidx].ToString();
                        if (strcheck.Length > 0)
                        {
                            selItem = new SelectedItemData();
                            selItem.CWKeyValue = CWKeyValue;
                            selItem.AttribObjectID = Convert.ToInt32(strcheck);
                            selItem.featuredescription = reader[descidx].ToString();
                            selItem.nodelevel = 1;
                            selItem.layername = ldef.layername;
                            SelectionSet.Add(selItem);
                        }
                    }
                }
                retval = SelectionSet.Count;
                connection.Close();
            }
            catch (Exception ex)
            {
                retval = 0;
            }
            return retval;
        }


        //------------------------------------------------------------------------------
        // Find Line Sections with GWUT By Level
        //------------------------------------------------------------------------------
        public bool find_Line_Section_Inspections(ref List<SelectedItemData> SelectionSet, LayerDef ldef)
        {
            // Get the Guided Wave Inspection Set ID
            int InspectionSetupID = getGenericTypeID("InspectionMethod:GWUT");
            // Build the query to search with
            string SQLStatement = "SELECT [Dbf Line Section].[Zone ID], [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point], [Dbf Line Section].[Line Section], [Dbf Line Section].[MapFeatureID], " +
                                  "[Dbf Inspection Range].Method FROM [Dbf Inspection Range] " +
                                  "INNER JOIN [Dbf Line Section] ON [Dbf Inspection Range].LineID = [Dbf Line Section].LineID " +
                                  "WHERE ((([Dbf Line Section].[Begin Point])>=[Dbf Inspection Range].[Begin Point]) AND (([Dbf Line Section].[End Point])<=[Dbf Inspection Range].[End Point]) AND (([Dbf Inspection Range].Method)=" + InspectionSetupID + "))";
            return ExecuteSQLtoSelectionSet(SQLStatement, ldef, ref SelectionSet);
        }


        //------------------------------------------------------------------------------
        // Find Lines with Licensed Material
        //------------------------------------------------------------------------------
        public bool find_Line_With_Licensed_Material(ref List<SelectedItemData> SelectionSet, LayerDef ldef)
        {
            // Build the query to search with
            string SQLStatement = "SELECT [Dbf Lines].LineID, [Dbf Lines].Line, [Dbf Lines].[Radiological Content], [Dbf Lines].MapFeatureID FROM [Dbf Lines] WHERE (([Dbf Lines].[Radiological Content])=True)";
            return ExecuteSQLtoSelectionSet(SQLStatement, ldef, ref SelectionSet);
        }


        //------------------------------------------------------------------------------
        // Execute a filter and populate the selected items list
        //------------------------------------------------------------------------------
        public bool ExecuteFiltertoSelectionSet(SQLFilterClass filter, LayerDefs layers, ref List<SelectedItemData> SelectionSet)
        {
            if (filter == null) return false;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                LayerDef ldef = new LayerDef();
                layers.GetLayerDef(filter.LayerIndex, ref ldef);
              
                // Set the sql command to the base query
                string cmdStr = ldef.BaseFilterQuery;
                // Add the conditions (WHERE)
                cmdStr += " WHERE (";
                int i;
                SQLCondition cond = new SQLCondition();
                for (i = 0; i <= (filter.ConditionCount - 1); i++)
                {
                    if (filter.GetCondition(i, ref cond))
                    {
                        if ((i > 0) && (cond.logicaloperator.Length > 0))
                            cmdStr += string.Format(" {0} ", cond.logicaloperator);
                        if (cond.fieldtype == "DBTYPE_WVARCHAR")
                            cmdStr += string.Format("([{0}] {1} '{2}')", cond.fieldname, cond.condition, cond.value);
                        else
                            cmdStr += string.Format("([{0}] {1} {2})", cond.fieldname, cond.condition, cond.value);
                    }
                }
                cmdStr += ")";
                int keyidx = 0; ;
                int descidx = 0;
                int attkeyidx = 0;
                SelectedItemData selItem; ;
                OleDbDataReader reader;
                OleDbCommand command = new OleDbCommand(cmdStr, connection);
                try
                {
                    reader = command.ExecuteReader();
                    keyidx = reader.GetOrdinal(ldef.GIS_CW_Index_Field);
                    descidx = reader.GetOrdinal(ldef.GIS_CW_Feature_Display_Field);
                    attkeyidx = reader.GetOrdinal(CW_XRef_Fieldname);


                    while (reader.Read())
                    {
                        if (reader[attkeyidx].ToString().Length > 0)
                        {
                            selItem = new SelectedItemData();
                            selItem.CWKeyValue = reader[keyidx].ToString();
                            selItem.AttribObjectID = Convert.ToInt32(reader[attkeyidx].ToString());
                            selItem.featuredescription = reader[descidx].ToString();
                            selItem.nodelevel = 1;
                            selItem.layername = ldef.layername;
                            SelectionSet.Add(selItem);
                        }
                    }
                    connection.Close();
                    return true;
                }
                catch (Exception ex)
                {
                    
                }
                connection.Close();
            }
            return false;
        }


        //------------------------------------------------------------------------------
        // Get Company Info
        //------------------------------------------------------------------------------
        private void GetCompany()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string queryString = "SELECT PlantName FROM [Dbf Company]";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        _companyname = reader[0].ToString();
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                }
                connection.Close();
            }
        }


        //-----------------------------------------------------------------------------------
        // Get Double Field Data from the Reader
        //-----------------------------------------------------------------------------------
        private bool GetFieldDataDouble(string DataFieldName, ref OleDbDataReader rdr, ref double dblvalue)
        {
            dblvalue = 0;
            try
            {
                int idx = rdr.GetOrdinal(DataFieldName);
                if (idx >= 0)
                {
                    string str = rdr[idx].ToString();
                    if (str.Length > 0)
                    {
                        dblvalue = Convert.ToDouble(str);
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            return false;
        }


        //-----------------------------------------------------------------------------------
        // Get Field Data from the Reader
        //-----------------------------------------------------------------------------------
        private string GetFieldDataStr(string DataFieldName, ref OleDbDataReader rdr)
        {
            string retval = "";
            try
            {
                int idx = rdr.GetOrdinal(DataFieldName);
                if (idx >= 0)
                    retval = rdr[idx].ToString();
            }
            finally
            {
            }
            return retval;
        }


        //-----------------------------------------------------------------------------------
        // Get Field Data from the Reader
        //-----------------------------------------------------------------------------------
        private int GetFieldDataInt(string DataFieldName, ref OleDbDataReader rdr)
        {
            int retval = -1;
            try
            {
                int idx = rdr.GetOrdinal(DataFieldName);
                if (idx >= 0)
                    retval = Convert.ToInt32(rdr[idx].ToString());
            }
            finally
            {
            }
            return retval;
        }

        //public static ESRI.ArcGIS.Display.IRgbColor MakeRGBColor(byte R, byte G, byte B)
        //{
        //    ESRI.ArcGIS.Display.RgbColor RgbClr = new ESRI.ArcGIS.Display.RgbColorClass();
        //    RgbClr.Red = R;
        //    RgbClr.Green = G;
        //    RgbClr.Blue = B;
        //    return RgbClr;
        //}

        //------------------------------------------------------------------------------
        // Loads the RiskRankingRange List
        //------------------------------------------------------------------------------
        public bool LoadRiskRankingRanges()
        {
            bool retval = false;
            string SQLStatement = "SELECT [Dbf RiskRange].LowSus, [Dbf RiskRange].LowMedSus, [Dbf RiskRange].MedSus, [Dbf RiskRange].MedHighSus, [Dbf RiskRange].HighSus, [Dbf RiskRange].LowCon, [Dbf RiskRange].LowMedCon, [Dbf RiskRange].MedCon, [Dbf RiskRange].MedHighCon, [Dbf RiskRange].HighCon, [Dbf RiskRange].LowRisk, [Dbf RiskRange].LowMedRisk, [Dbf RiskRange].MedRisk, [Dbf RiskRange].MedHighRisk, [Dbf RiskRange].HighRisk FROM [Dbf RiskRange]";
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            connection.Open();
            // Set the sql command to the base query
            string cmdStr = SQLStatement;
            OleDbDataReader reader;
            OleDbCommand command = new OleDbCommand(cmdStr, connection);
            RiskRankingRange RRR;
            int val = 0;
            int val2 = 0;
            try
            {
                RiskRankingRanges.Clear();
                RiskRankingRangesSusc.Clear();
                RiskRankingRangesCons.Clear();
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    // Susceptibility Ranges
                    // Low
                    val = GetFieldDataInt("LowSus", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(0, val,0,255,0);
                    RiskRankingRangesSusc.Add(RRR);
                    // Low - Med
                    val2 = val + 1;
                    val = GetFieldDataInt("LowMedSus", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2,val,255,255,128);
                    RiskRankingRangesSusc.Add(RRR);
                    // Med
                    val2 = val + 1;
                    val = GetFieldDataInt("MedSus", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val,255,192,128);
                    RiskRankingRangesSusc.Add(RRR);
                    // Med High
                    val2 = val + 1;
                    val = GetFieldDataInt("MedHighSus", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val,255,128,255);
                    RiskRankingRangesSusc.Add(RRR);
                    // High
                    val2 = val + 1;
                    val = GetFieldDataInt("HighSus", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val,255,0,0);
                    RiskRankingRangesSusc.Add(RRR);

                    // Consequence Ranges
                    // Low
                    val = GetFieldDataInt("LowCon", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(0, val, 0,255,0);
                    RiskRankingRangesCons.Add(RRR);
                    // Low - Med
                    val2 = val + 1;
                    val = GetFieldDataInt("LowMedCon", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val, 255, 255, 128);
                    RiskRankingRangesCons.Add(RRR);
                    // Med
                    val2 = val + 1;
                    val = GetFieldDataInt("MedCon", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val, 255, 192, 128);
                    RiskRankingRangesCons.Add(RRR);
                    // Med High
                    val2 = val + 1;
                    val = GetFieldDataInt("MedHighCon", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val, 255,128,255);
                    RiskRankingRangesCons.Add(RRR);
                    // High
                    val2 = val + 1;
                    val = GetFieldDataInt("HighCon", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val,255,0,0);
                    RiskRankingRangesCons.Add(RRR);


                    // Overall Ranges
                    // Low
                    val = GetFieldDataInt("LowRisk", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val,0,255,0);
                    RiskRankingRanges.Add(RRR);
                    // Low - Med
                    val2 = val + 1;
                    val = GetFieldDataInt("LowMedRisk", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val, 255,255,128);
                    RiskRankingRanges.Add(RRR);
                    // Med
                    val2 = val + 1;
                    val = GetFieldDataInt("MedRisk", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val, 255,192,128);
                    RiskRankingRanges.Add(RRR);
                    // Med High
                    val2 = val + 1;
                    val = GetFieldDataInt("MedHighRisk", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val,255,128,255);
                    RiskRankingRanges.Add(RRR);
                    // High
                    val2 = val + 1;
                    val = GetFieldDataInt("HighRisk", ref reader);
                    if (val <= 0) return false;
                    RRR = new RiskRankingRange(val2, val,255,0,0);
                    RiskRankingRanges.Add(RRR);
                }
                connection.Close();
            }
            catch (Exception ex)
            {
            }
            connection.Close();
            return retval;
        }



        //------------------------------------------------------------------------------
        // Get Items Needing Update from CompWorks
        // Returns a string with the characters representing the types of updates found
        //------------------------------------------------------------------------------
        public bool linesectionCheckForUpdates(string flagChar)
        {
            bool retval = false;

            // Check the "Risk Ranking" [R]
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string cmdStr = "SELECT [Dbf Line Section].[Zone ID], [Dbf Line Section].[Line Section], [Dbf Line Section].SusRanking, [Dbf Line Section].ConRanking, [Dbf Line Section].[Overall Ranking], [Dbf Line Section].UpdateFlags FROM [Dbf Line Section] WHERE ((([Dbf Line Section].UpdateFlags) Like '%" + flagChar + "%'))";
                try
                {
                    OleDbCommand command = new OleDbCommand(cmdStr, connection);
                    if (command != null)
                    {
                        OleDbDataReader reader = command.ExecuteReader();
                        retval = (reader.HasRows);
                        reader.Close();
                    }
                    connection.Close();
                }
                catch (Exception ex)
                {
                    return retval;
                }
            }
            return retval;
        }



        //------------------------------------------------------------------------------
        // Set Line Section Record UpdateFlags value
        //------------------------------------------------------------------------------
        public bool linesectionSetUpdateFlag(int CWID, string flagStr)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = string.Format("UPDATE [Dbf Line Section] SET [Dbf Line Section].UpdateFlags = '{0}' WHERE (([Dbf Line Section].[Zone ID]={1}))", flagStr, CWID);
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }


        //------------------------------------------------------------------------------
        // Set Layer Name
        //------------------------------------------------------------------------------
        public bool SetLayerName(int CWSystemUse, string newLayerName)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = string.Format("UPDATE [GIS_Layers] SET [GIS_Layers].GIS_Layer_Name = '{0}' WHERE (([GIS_Layers].GIS_Layer_SystemUse={1}))", newLayerName, CWSystemUse);
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }

        //------------------------------------------------------------------------------
        // Set Layer CW Key Field
        //------------------------------------------------------------------------------
        public bool SetLayerCWKeyField(int CWSystemUse, string newKeyFieldName)
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = string.Format("UPDATE [GIS_Layers] SET [GIS_Layers].GIS_Attribute_Index_Field = '{0}' WHERE (([GIS_Layers].GIS_Layer_SystemUse={1}))", newKeyFieldName, CWSystemUse);
                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
                return true;
            }
            return false;
        }



        //------------------------------------------------------------------------------
        // Get Items Needing Update from CompWorks
        // Returns a string with the characters representing the types of updates found
        //------------------------------------------------------------------------------
        public bool linesectionRemoveUpdateFlag(string flagChar,int zoneid)
        {
            bool retval = false;
            if (zoneid <= 0) return retval;
            // Check the "Risk Ranking" [R]
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string cmdStr = "SELECT [Dbf Line Section].[Zone ID], [Dbf Line Section].UpdateFlags FROM [Dbf Line Section] WHERE ((([Dbf Line Section].[Zone ID])=" + zoneid.ToString() + "))";
                try
                {
                    OleDbCommand command = new OleDbCommand(cmdStr, connection);
                    if (command != null)
                    {
                        OleDbDataReader reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            string FlagStr = GetFieldDataStr("UpdateFlags", ref reader);
                            //if (string.Compare(FlagStr, flagChar) == 0)
                            //    FlagStr = "";
                            //else
                            if (FlagStr.IndexOf(flagChar) >= 0)
                            {
                                FlagStr = FlagStr.Remove(FlagStr.IndexOf(flagChar), 1);
                                linesectionSetUpdateFlag(zoneid, FlagStr);
                            }
                        }
                        reader.Close();
                    }
                    connection.Close();
                }
                catch (Exception ex)
                {
                    return retval;
                }
            }
            return retval;
        }



        //------------------------------------------------------------------------------
        // class constructor
        //------------------------------------------------------------------------------
        public CompWorksDBConnect(string DBFileName, string SystemFileName)
        {
            _dbfilename = DBFileName;
            _systemfilename = SystemFileName;
            _connectionstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + DBFileName + "; Jet OLEDB:System Database=" + SystemFileName + "; User ID=developer; Password=waiheke";
            // Create the configuration object
            Configuration = new Configuration();
            Configuration.FilePath = System.IO.Path.ChangeExtension(Application.ExecutablePath, ".cfg");
            Configuration.LoadConfigFile();
            GetCompany();
            LoadRiskRankingRanges();
        }

    }
}

