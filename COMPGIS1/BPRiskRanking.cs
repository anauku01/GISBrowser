using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace COMPGIS1
{
    public class BPRiskRankingItem
    {
        public string RiskRankingDesc = "";
        public int RiskRankingValue = 0;
        public string RiskRankingFieldNameLS = "";
        public string RiskRankingFieldNameRR = "";

        public BPRiskRankingItem(string Desc, int Value, string FieldNameLS, string FieldNameRR)
        {
            RiskRankingDesc = Desc;
            RiskRankingValue = Value;
            RiskRankingFieldNameLS = FieldNameLS;
            RiskRankingFieldNameRR = FieldNameRR;
        }
    }

    // ------------------------------------------------------------
    // BPRiskRankingValues
    // ------------------------------------------------------------
    public class BPRiskRankingValues
    {
        private List<BPRiskRankingItem> _items;

        public int Count
        {
            get { return _items.Count; }
            set { }
        }

        // Constructor
        public BPRiskRankingValues()
        {
            _items = new List<BPRiskRankingItem>();
        }

        // Clears the items list
        public void Clear()
        {
            _items.Clear();
        }

        // Adds a RR Item to the List
        public bool AddItem(string RRDesc, int RRValue, string RRFieldNameLS, string RRFieldName)
        {
            if (RRDesc.Length == 0) return false;
            BPRiskRankingItem newitem = new BPRiskRankingItem(RRDesc, RRValue, RRFieldNameLS, RRFieldName);
            if (!(newitem == null))
            {
                _items.Add(newitem);
                return true;
            }
            return false;
        }


        // Returns a specific RR Item
        public bool GetItem(int index, ref BPRiskRankingItem RRItem)
        {
            RRItem.RiskRankingDesc = "";
            RRItem.RiskRankingValue = 0;
            if ((index < 0) || (index > (_items.Count - 1))) return false;
            RRItem = _items[index];
            return true;
        }
    }





    // ------------------------------------------------------------
    // ------------------------------------------------------------
    // ------------------------------------------------------------
    // BPSuscRiskRankingValues
    // ------------------------------------------------------------
    // ------------------------------------------------------------
    // ------------------------------------------------------------
    public class BPSusRiskRankingValues
    {


        private BPRiskRankingValues _items;
        private long _lineid = 0;
        private long _beginpt = 0;
        private long _endpt = 0;
        public string _connectionstring = "";

        public BPRiskRankingValues RrValues
        {
            get { return _items; }
            set { }
        }

        // Constructor
        public BPSusRiskRankingValues(long LineID, long BeginPoint, long EndPoint, string cs)
        {
            if ((LineID > 0) && (BeginPoint >= 0) && (EndPoint >= 0) && (cs.Length > 0))
            {
                _connectionstring = cs;
                _lineid = LineID;
                _beginpt = BeginPoint;
                _endpt = EndPoint;
                _items = new BPRiskRankingValues();
                LoadTheStructures();
            }
        }

        // Constructor 2
        public BPSusRiskRankingValues(string cs)
        {
            if (cs.Length > 0)
            {
                _connectionstring = cs;
                _items = new BPRiskRankingValues();
                LoadAllFactors();
            }
        }


        //-----------------------------------------------------------------------------------
        // See if the field exists
        //-----------------------------------------------------------------------------------
        public bool ColumnExists(IDataReader reader, string columnName)
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (string.Compare(columnName, reader.GetName(i)) == 0)
                {
                    return true;
                }
            }

            return false;
        }


        //-----------------------------------------------------------------------------------
        // Get BOOLEAN Field Data from the Reader
        //-----------------------------------------------------------------------------------
        private bool GetFieldDataBool(string DataFieldName, ref OleDbDataReader rdr, bool defvalue)
        {
            bool retval = defvalue;
            try
            {
                int idx = rdr.GetOrdinal(DataFieldName);
                if (idx >= 0)
                {
                    string str = rdr[idx].ToString();
                    if (str.Length > 0)
                        retval = (string.Compare(str, "true", true) == 0);
                }
            }
            finally
            {
            }
            return retval;
        }

        //-----------------------------------------------------------------------------------
        // Get STRING Field Data from the Reader
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
        // Get INT Field Data from the Reader
        //-----------------------------------------------------------------------------------
        private int GetFieldDataInt(string DataFieldName, ref OleDbDataReader rdr)
        {
            int retval = -1;
            try
            {
                int idx = rdr.GetOrdinal(DataFieldName);
                if (idx >= 0)
                {
                    string ssss = rdr[idx].ToString();
                    if (ssss.Length > 0)
                        retval = Convert.ToInt32(rdr[idx].ToString());
                }
            }
            finally
            {
            }
            return retval;
        }


        // Load All Factors in to the structures
        private void LoadAllFactors()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string queryString = "Select * from [Dbf Setup - RiskRanking Sus]";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    double idx = 0;
                    OleDbDataReader RRSus = command.ExecuteReader();
                    if (RRSus.Read())
                    {
                        _items.Clear();
                        _items.AddItem("Piping Under Roadway", GetFieldDataInt("VUnderRoad", ref RRSus), "Under_Road",
                                       "VUnderRoad");
                        _items.AddItem("Piping Under Building", GetFieldDataInt("VUnderBuilding", ref RRSus),
                                       "Under_Bldg", "VUnderBuilding");
                        _items.AddItem("Piping Under Railroad Tracks", GetFieldDataInt("VUnderRR", ref RRSus),
                                       "Under_RR", "VUnderRR");
                        _items.AddItem("Under Transmission Line", GetFieldDataInt("VHighVolt", ref RRSus),
                                       "Under_Trans_Line", "VHighVolt");
                        _items.AddItem("Above or Under Water or Canal", GetFieldDataInt("VBridgeCanal", ref RRSus),
                                       "Over_Under_River", "VBridgeCanal");
                        _items.AddItem("Pipe Replaced with New Pipe", GetFieldDataInt("VNewPipe", ref RRSus), "Replaced",
                                       "VNewPipe");
                        _items.AddItem("No Pipe Wall Inspection", GetFieldDataInt("VInspNotDone", ref RRSus),
                                       "Inspected", "VInspNotDone");
                        _items.AddItem("Change of Earthen Material Fill", GetFieldDataInt("VEarthFillChg", ref RRSus),
                                       "Earthen_Fill_Material_Change", "VEarthFillChg");
                        _items.AddItem("Piping Not Coated", GetFieldDataInt("VNotCoated", ref RRSus), "NotCoated",
                                       "VNotCoated");
                        _items.AddItem("Change in Pipe Material (Less Noble Material)",
                                       GetFieldDataInt("VMtlChange", ref RRSus), "Mtl_Chg", "VMtlChange");
                        _items.AddItem("Connected to Building/Concrete", GetFieldDataInt("VInOutBldg", ref RRSus),
                                       "In_To_Out_Of_Bldg", "VInOutBldg");
                        _items.AddItem("Below Grade Tees or Junctions", GetFieldDataInt("VUndGrT", ref RRSus),
                                       "Underground_Tee", "VUndGrT");
                        _items.AddItem("Evidence of Ground Settlement", GetFieldDataInt("VGroundSetYes", ref RRSus),
                                       "GroundSettlement", "VGroundSetYes");
                        _items.AddItem("Metallic Pipes or Ducts in Area", GetFieldDataInt("VPipeInArea", ref RRSus),
                                       "PipesInArea", "VPipeInArea");
                        _items.AddItem("Buried Steam Line in Area", GetFieldDataInt("VSteamLineInArea", ref RRSus),
                                       "SteamLine", "VSteamLineInArea");
                        _items.AddItem("Change in Pipe Wall Thickness", GetFieldDataInt("VThinPipe", ref RRSus),
                                       "WallThinner", "VThinPipe");
                        _items.AddItem("Potential for Pressure Transients", GetFieldDataInt("VPotential", ref RRSus),
                                       "Potential", "VPotential");
                        _items.AddItem("Fluid Temperature >200 degrees", GetFieldDataInt("VTempOver200", ref RRSus),
                                       "Temp200", "VTempOver200");
                        _items.AddItem("Contains Corrosive Fluid", GetFieldDataInt("VCorrFluid", ref RRSus),
                                       "CorrosiveFluid", "VCorrFluid");
                        _items.AddItem("Pipe Not Subjected to Chemical/Biocide Additions",
                                       GetFieldDataInt("VChemAddNo", ref RRSus), "NoChemicalAdditions", "VChemAddNo");
                        _items.AddItem("Piping Above Groundwater Level",
                                       GetFieldDataInt("VPipeAboveGrndWater", ref RRSus), "Above_Groundwater_Level",
                                       "VPipeAboveGrndWater");
                        _items.AddItem("No Coating Inspection Performed", GetFieldDataInt("VCoatInspNP", ref RRSus),
                                       "No Coating Inspection Performed", "VCoatInspNP");
                        _items.AddItem("Coating Degradation Identified", GetFieldDataInt("VCoatHoliday", ref RRSus),
                                       "Coating Degradation Identified", "VCoatHoliday");
                        _items.AddItem("Within 10 ft of Transmission Line Footer",
                                       GetFieldDataInt("VConcFooter", ref RRSus),
                                       "Within 10 ft Transmission Line Footer", "VConcFooter");
                        _items.AddItem("Detrimental Flow Erosion/Corrosion Wall Loss Contributor",
                                       GetFieldDataInt("VInternalECContYes", ref RRSus), "Internal Erosion Corrosion",
                                       "VInternalECContYes");
                        _items.AddItem("Soil Characteristics Unknown", GetFieldDataInt("VSoilVerifiedNo", ref RRSus),
                                       "Soil Characteristics Unknown", "VSoilVerifiedNo");
                        _items.AddItem("Recorded Transients Not corrected to Prevent Recurrence",
                                       GetFieldDataInt("OperTransYes", ref RRSus), "Recorded Transient Not Corrected",
                                       "OperTransYes");
                        _items.AddItem("Pipe Wall Degradation or Loss Identified",
                                       GetFieldDataInt("VInspLoss", ref RRSus), "Pipe Wall Degradation", "VInspLoss");
                        _items.AddItem("Cathodic Protection Not Functional", GetFieldDataInt("VCPNonFunc", ref RRSus),
                                       "RectifierOperational", "VCPNonFunc");

                        _items.AddItem("Vertical Runs of Pipe", 1, "Pipe", "VVert");
                        _items.AddItem("Piping Elevation Changes (Slopes up or down)", 3, "Pipe", "VElevChg");

                        _items.AddItem("1 to 2 Leaks in Zone", 1, "Leak_History","V12Leak");
                        _items.AddItem("3 to 4 Leaks in Zone", 2 , "Leak_History","V34Leak");

                        _items.AddItem("Coating > 10 Years but <30 Years", 1,"PipeAge", "V1030");
                        _items.AddItem("Coating > 30 Years", 2, "PipeAge", "V30Up");

                        _items.AddItem("Susceptibility Engineering Judgement", 0, "Susceptibility Eng Judgment Value",
                                       "Susceptibility Eng Judgment Value");
                    }
                    RRSus.Close();
                }
                catch (Exception ex)
                {
                    if (ex.Message.Length > 0)
                        return;
                }
                connection.Close();
            }
        }


        // Load the Values
        private void LoadTheStructures()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string queryString = "Select * from [Dbf Setup - RiskRanking Sus]";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    double idx = 0;
                    OleDbDataReader RRSus = command.ExecuteReader();
                    if (RRSus.Read())
                    {
                        queryString = "Select * from [Dbf Susceptibility LS] WHERE LineID= " + _lineid +
                                      " AND ([Begin Point] <= " + _beginpt + " AND [End Point] >= " + _endpt + ")";
                        command = new OleDbCommand(queryString, connection);
                        OleDbDataReader LSSus = command.ExecuteReader();
                        if (LSSus.Read())
                        {
                            _items.Clear();
                            if (GetFieldDataBool("Under_Road", ref LSSus, false))
                                _items.AddItem("Piping Under Roadway", GetFieldDataInt("VUnderRoad", ref RRSus),
                                               "Under_Road", "VUnderRoad");
                            if (GetFieldDataBool("Under_Bldg", ref LSSus, false))
                                _items.AddItem("Piping Under Building", GetFieldDataInt("VUnderBuilding", ref RRSus),
                                               "Under_Bldg", "VUnderBuilding");
                            if (GetFieldDataBool("Under_RR", ref LSSus, false))
                                _items.AddItem("Piping Under Railroad Tracks", GetFieldDataInt("VUnderRR", ref RRSus),
                                               "Under_RR", "VUnderRR");
                            if (GetFieldDataBool("Under_Trans_Line", ref LSSus, false))
                                _items.AddItem("Under Transmission Line", GetFieldDataInt("VHighVolt", ref RRSus),
                                               "Under_Trans_Line", "VHighVolt");
                            if (GetFieldDataBool("Over_Under_River", ref LSSus, false))
                                _items.AddItem("Above or Under Water or Canal",
                                               GetFieldDataInt("VBridgeCanal", ref RRSus), "Over_Under_River",
                                               "VBridgeCanal");
                            if (GetFieldDataBool("Replaced", ref LSSus, false))
                                _items.AddItem("Pipe Replaced with New Pipe", GetFieldDataInt("VNewPipe", ref RRSus),
                                               "Replaced", "VNewPipe");
                            if (GetFieldDataBool("Inspected", ref LSSus, false))
                                _items.AddItem("No Pipe Wall Inspection", GetFieldDataInt("VInspNotDone", ref RRSus),
                                               "Inspected", "VInspNotDone");
                            if (GetFieldDataBool("Earthen_Fill_Material_Change", ref LSSus, false))
                                _items.AddItem("Change of Earthen Material Fill",
                                               GetFieldDataInt("VEarthFillChg", ref RRSus),
                                               "Earthen_Fill_Material_Change", "VEarthFillChg");
                            if (GetFieldDataBool("NotCoated", ref LSSus, false))
                                _items.AddItem("Piping Not Coated", GetFieldDataInt("VNotCoated", ref RRSus),
                                               "NotCoated", "VNotCoated");
                            if (GetFieldDataBool("Mtl_Chg", ref LSSus, false))
                                _items.AddItem("Change in Pipe Material (Less Noble Material)",
                                               GetFieldDataInt("VMtlChange", ref RRSus), "Mtl_Chg", "VMtlChange");
                            if (GetFieldDataBool("In_To_Out_Of_Bldg", ref LSSus, false))
                                _items.AddItem("Connected to Building/Concrete",
                                               GetFieldDataInt("VInOutBldg", ref RRSus), "In_To_Out_Of_Bldg",
                                               "VInOutBldg");
                            if (GetFieldDataBool("Underground_Tee", ref LSSus, false))
                                _items.AddItem("Below Grade Tees or Junctions", GetFieldDataInt("VUndGrT", ref RRSus),
                                               "Underground_Tee", "VUndGrT");
                            if (GetFieldDataBool("GroundSettlement", ref LSSus, false))
                                _items.AddItem("Evidence of Ground Settlement",
                                               GetFieldDataInt("VGroundSetYes", ref RRSus), "GroundSettlement",
                                               "VGroundSetYes");
                            if (GetFieldDataBool("PipesInArea", ref LSSus, false))
                                _items.AddItem("Metallic Pipes or Ducts in Area",
                                               GetFieldDataInt("VPipeInArea", ref RRSus), "PipesInArea", "VPipeInArea");
                            if (GetFieldDataBool("SteamLine", ref LSSus, false))
                                _items.AddItem("Buried Steam Line in Area",
                                               GetFieldDataInt("VSteamLineInArea", ref RRSus), "SteamLine",
                                               "VSteamLineInArea");
                            if (GetFieldDataBool("WallThinner", ref LSSus, false))
                                _items.AddItem("Change in Pipe Wall Thickness", GetFieldDataInt("VThinPipe", ref RRSus),
                                               "WallThinner", "VThinPipe");
                            if (GetFieldDataBool("Potential", ref LSSus, false))
                                _items.AddItem("Potential for Pressure Transients",
                                               GetFieldDataInt("VPotential", ref RRSus), "Potential", "VPotential");
                            if (GetFieldDataBool("Temp200", ref LSSus, false))
                                _items.AddItem("Fluid Temperature >200 degrees",
                                               GetFieldDataInt("VTempOver200", ref RRSus), "Temp200", "VTempOver200");
                            if (GetFieldDataBool("CorrosiveFluid", ref LSSus, false))
                                _items.AddItem("Contains Corrosive Fluid", GetFieldDataInt("VCorrFluid", ref RRSus),
                                               "CorrosiveFluid", "VCorrFluid");
                            if (GetFieldDataBool("NoChemicalAdditions", ref LSSus, false))
                                _items.AddItem("Pipe Not Subjected to Chemical/Biocide Additions",
                                               GetFieldDataInt("VChemAddNo", ref RRSus), "NoChemicalAdditions",
                                               "VChemAddNo");
                            if (GetFieldDataBool("Above_Groundwater_Level", ref LSSus, false))
                                _items.AddItem("Piping Above Groundwater Level",
                                               GetFieldDataInt("VPipeAboveGrndWater", ref RRSus),
                                               "Above_Groundwater_Level", "VPipeAboveGrndWater");
                            if (GetFieldDataBool("No Coating Inspection Performed", ref LSSus, false))
                                _items.AddItem("No Coating Inspection Performed",
                                               GetFieldDataInt("VCoatInspNP", ref RRSus),
                                               "No Coating Inspection Performed", "VCoatInspNP");
                            if (GetFieldDataBool("Coating Degradation Identified", ref LSSus, false))
                                _items.AddItem("Coating Degradation Identified",
                                               GetFieldDataInt("VCoatHoliday", ref RRSus),
                                               "Coating Degradation Identified", "VCoatHoliday");
                            if (GetFieldDataBool("Within 10 ft Transmission Line Footer", ref LSSus, false))
                                _items.AddItem("Within 10 ft of Transmission Line Footer",
                                               GetFieldDataInt("VConcFooter", ref RRSus),
                                               "Within 10 ft Transmission Line Footer", "VConcFooter");
                            if (GetFieldDataBool("Internal Erosion Corrosion", ref LSSus, false))
                                _items.AddItem("Detrimental Flow Erosion/Corrosion Wall Loss Contributor",
                                               GetFieldDataInt("VInternalECContYes", ref RRSus),
                                               "Internal Erosion Corrosion", "VInternalECContYes");
                            if (GetFieldDataBool("Soil Characteristics Unknown", ref LSSus, false))
                                _items.AddItem("Soil Characteristics Unknown",
                                               GetFieldDataInt("VSoilVerifiedNo", ref RRSus),
                                               "Soil Characteristics Unknown", "VSoilVerifiedNo");
                            if (GetFieldDataBool("Recorded Transient Not Corrected", ref LSSus, false))
                                _items.AddItem("Recorded Transients Not corrected to Prevent Recurrence",
                                               GetFieldDataInt("OperTransYes", ref RRSus),
                                               "Recorded Transient Not Corrected", "OperTransYes");
                            if (GetFieldDataBool("Pipe Wall Degradation", ref LSSus, false))
                                _items.AddItem("Pipe Wall Degradation or Loss Identified",
                                               GetFieldDataInt("VInspLoss", ref RRSus), "Pipe Wall Degradation",
                                               "VInspLoss");
                            if (GetFieldDataBool("RectifierOperational", ref LSSus, false))
                                _items.AddItem("Cathodic Protection Not Functional",
                                               GetFieldDataInt("VCPNonFunc", ref RRSus), "RectifierOperational",
                                               "VCPNonFunc");
                            // Engineering Judgement
                            int n = GetFieldDataInt("Susceptibility Eng Judgment Value", ref LSSus);
                            if (n > 0)
                                _items.AddItem("Susceptibility Engineering Judgement",
                                               GetFieldDataInt("Susceptibility Eng Judgment Value", ref LSSus),
                                               "Susceptibility Eng Judgment Value", "Susceptibility Eng Judgment Value");
                            // Vertical / Horizontal Runs of Pipe
                            n = GetFieldDataInt("Pipe", ref LSSus);
                            if (n == 1)
                            {
                                _items.AddItem("Vertical Runs of Pipe", GetFieldDataInt("VVert", ref RRSus), "Pipe", "VVert");

                            }
                            else
                            {
                                if (n == 3)
                                {
                                    _items.AddItem("Piping Elevation Changes (Slopes up or down)", GetFieldDataInt("VElevChg", ref RRSus), "Pipe", "VElevChg");
                                }
                            }
                            // Leak History is a number where (1 = 1 to 2 leaks; 3 = 3 to 4 leaks)
                            n = GetFieldDataInt("Leak_History", ref LSSus);
                            if (n == 1)
                            {
                                _items.AddItem("1 to 2 Leaks in Zone", GetFieldDataInt("V12Leak", ref RRSus), "Leak_History", "V12Leak");
                            }
                            else
                            {
                                if (n == 2)
                                {
                                    _items.AddItem("3 to 4 Leaks in Zone", GetFieldDataInt("V34Leak", ref RRSus), "Leak_History", "V34Leak");
                                }
                            }
                            // PipeAge is a number where (1 = 10 to 30yrs; 2 = 30 and up)
                            n = GetFieldDataInt("PipeAge", ref LSSus);
                            if (n == 1)
                            {
                                _items.AddItem("Coating > 10 Years but <30 Years", GetFieldDataInt("V1030", ref RRSus), "PipeAge", "V1030");
                            }
                            else
                            {
                                if (n == 2)
                                {
                                    _items.AddItem("Coating > 30 Years", GetFieldDataInt("V30Up", ref RRSus), "PipeAge", "V30Up");
                                }
                            }
                            // Done
                            LSSus.Close();
                        }
                    }
                    RRSus.Close();
                }
                catch (Exception ex)
                {
                    if (ex.Message.Length > 0)
                        return;
                }
                connection.Close();
            }
        }
    }







    // ------------------------------------------------------------
    // ------------------------------------------------------------
    // ------------------------------------------------------------
    // BPConRiskRankingValues
    // ------------------------------------------------------------
    // ------------------------------------------------------------
    // ------------------------------------------------------------
    public class BPConRiskRankingValues
    {

        private BPRiskRankingValues _items;
        private long _lineid = 0;
        private long _beginpt = 0;
        private long _endpt = 0;
        public string _connectionstring = "";

        public BPRiskRankingValues RrValues
        {
            get { return _items; }
            set { }
        }

        // Constructor
        // This constructor loads the risk ranking factors and values for a particular line section
        public BPConRiskRankingValues(long LineID, long BeginPoint, long EndPoint, string cs)
        {
            if ((LineID > 0) && (BeginPoint >= 0) && (EndPoint >= 0) && (cs.Length > 0))
            {
                _connectionstring = cs;
                _lineid = LineID;
                _beginpt = BeginPoint;
                _endpt = EndPoint;
                _items = new BPRiskRankingValues();
                LoadTheStructures();
            }
        }

        // Constructor 2
        // This constructor loads all the risk ranking factors in to the list
        public BPConRiskRankingValues(string cs)
        {
            if (cs.Length > 0)
            {
                _connectionstring = cs;
                _items = new BPRiskRankingValues();
                LoadAllFactors();
            }
        }


        //-----------------------------------------------------------------------------------
        // See if the field exists
        //-----------------------------------------------------------------------------------
        public bool ColumnExists(IDataReader reader, string columnName)
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (string.Compare(columnName, reader.GetName(i)) == 0)
                {
                    return true;
                }
            }

            return false;
        }


        //-----------------------------------------------------------------------------------
        // Get BOOLEAN Field Data from the Reader
        //-----------------------------------------------------------------------------------
        private bool GetFieldDataBool(string DataFieldName, ref OleDbDataReader rdr, bool defvalue)
        {
            bool retval = defvalue;
            try
            {
                int idx = rdr.GetOrdinal(DataFieldName);
                if (idx >= 0)
                    retval = (string.Compare(rdr[idx].ToString(), "true", true) == 0);
            }
            finally
            {
            }
            return retval;
        }

        //-----------------------------------------------------------------------------------
        // Get STRING Field Data from the Reader
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
        // Get INT Field Data from the Reader
        //-----------------------------------------------------------------------------------
        private int GetFieldDataInt(string DataFieldName, ref OleDbDataReader rdr)
        {
            int retval = -1;
            try
            {
                int idx = rdr.GetOrdinal(DataFieldName);
                if (idx >= 0)
                {
                    retval = Convert.ToInt32(rdr[idx].ToString());
                }
            }
            finally
            {
            }
            return retval;
        }


        // Load the Values
        private void LoadTheStructures()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string queryString = "Select * from [Dbf Setup - RiskRanking Con]";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    double idx = 0;
                    OleDbDataReader RRSus = command.ExecuteReader();
                    if (RRSus.Read())
                    {
                        queryString = "Select * from [Dbf Consequence] WHERE LineID= " + _lineid +
                                      " AND ([Begin Point] <= " + _beginpt + " AND [End Point] >= " + _endpt + ")";
                        command = new OleDbCommand(queryString, connection);
                        OleDbDataReader LSCon = command.ExecuteReader();
                        if (LSCon.Read())
                        {
                            _items.Clear();
                            if (GetFieldDataBool("Licensing", ref LSCon, false))
                                _items.AddItem("Piping Has Specific Licensing Commitments",
                                               GetFieldDataInt("VLicensing", ref RRSus), "Licensing", "VLicensing");
                            if (GetFieldDataBool("Repair1Million", ref LSCon, false))
                                _items.AddItem("Repair > 1 Million Dollars",
                                               GetFieldDataInt("VRepair1Million", ref RRSus), "Repair1Million",
                                               "VRepair1Million");
                            if (GetFieldDataBool("Radiological", ref LSCon, false))
                                _items.AddItem("Piping Contains Radiological Fluid",
                                               GetFieldDataInt("VRadiological", ref RRSus), "Radiological",
                                               "VRadiological");
                            if (GetFieldDataBool("EPA", ref LSCon, false))
                                _items.AddItem("Piping Contains EPA Sensitive Fluid", GetFieldDataInt("VEPA", ref RRSus),
                                               "EPA", "VEPA");
                            if (GetFieldDataBool("UnderBldg2", ref LSCon, false))
                                _items.AddItem("Piping Contains EPA Sensitive Fluid",
                                               GetFieldDataInt("VUnderBuilding2", ref RRSus), "UnderBldg2",
                                               "VUnderBuilding2");
                            if (GetFieldDataBool("tritium", ref LSCon, false))
                                _items.AddItem("Process Pipe Carries Tritium", GetFieldDataInt("Vtritium", ref RRSus),
                                               "tritium", "Vtritium");
                            if (GetFieldDataBool("RadWaste", ref LSCon, false))
                                _items.AddItem("Process Piping Carries Rad Waste",
                                               GetFieldDataInt("VRadWaste", ref RRSus), "RadWaste", "VRadWaste");
                            if (GetFieldDataBool("NSIAC", ref LSCon, false))
                                _items.AddItem("Line Section is Governed by NSIAC", GetFieldDataInt("VNSIAC", ref RRSus),
                                               "NSIAC", "VNSIAC");
                            if (GetFieldDataBool("SafetyRelated", ref LSCon, false))
                                _items.AddItem("Safety Related Service Water",
                                               GetFieldDataInt("VSafetyRelated", ref RRSus), "SafetyRelated",
                                               "VSafetyRelated");

                            // LCO
                            int n = GetFieldDataInt("LCO", ref LSCon);
                            switch (n)
                            {
                                case 1:
                                    _items.AddItem("LCO <72 Hours", GetFieldDataInt("VLCOLT72", ref RRSus), "LCO",
                                                   "VLCOLT72");
                                    break;
                                case 2:
                                    _items.AddItem("LCO 72 Hours", GetFieldDataInt("VLCO72", ref RRSus), "LCO", "VLCO72");
                                    break;
                                case 3:
                                    _items.AddItem("LCO 7 Day", GetFieldDataInt("VLCO7", ref RRSus), "LCO", "VLCO7");
                                    break;
                                case 4:
                                    _items.AddItem("LCO 14 Day", GetFieldDataInt("VLCO14", ref RRSus), "LCO", "VLCO14");
                                    break;
                                case 5:
                                    _items.AddItem("LCO 30 Day", GetFieldDataInt("VLCO30", ref RRSus), "LCO", "VLCO30");
                                    break;
                            }
                            // Engineering Judgement
                            n = GetFieldDataInt("Consequence Eng Judgment Value", ref LSCon);
                            if (n > 0)
                                _items.AddItem("Consequence Eng Judgment",
                                               GetFieldDataInt("Consequence Eng Judgment Value", ref LSCon),
                                               "Consequence Eng Judgment Value", "Consequence Eng Judgment Value");
                            // Done
                            LSCon.Close();
                        }
                    }
                    RRSus.Close();
                }
                catch (Exception ex)
                {
                    if (ex.Message.Length > 0)
                        return;
                }
                connection.Close();
            }
        }


        // Load the Values
        private void LoadAllFactors()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string queryString = "Select * from [Dbf Setup - RiskRanking Con]";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    double idx = 0;
                    OleDbDataReader RRSus = command.ExecuteReader();
                    if (RRSus.Read())
                    {
                        _items.AddItem("Piping Has Specific Licensing Commitments",
                                       GetFieldDataInt("VLicensing", ref RRSus), "Licensing", "VLicensing");
                        _items.AddItem("Repair > 1 Million Dollars", GetFieldDataInt("VRepair1Million", ref RRSus),
                                       "Repair1Million", "VRepair1Million");
                        _items.AddItem("Piping Contains Radiological Fluid", GetFieldDataInt("VRadiological", ref RRSus),
                                       "Radiological", "VRadiological");
                        _items.AddItem("Piping Contains EPA Sensitive Fluid", GetFieldDataInt("VEPA", ref RRSus), "EPA",
                                       "VEPA");
                        _items.AddItem("Undermine Building or Supports of Other Systems",
                                       GetFieldDataInt("VUnderBuilding2", ref RRSus), "UnderBldg2", "VUnderBuilding2");
                        _items.AddItem("Process Pipe Carries Tritium", GetFieldDataInt("Vtritium", ref RRSus), "tritium",
                                       "Vtritium");
                        _items.AddItem("Process Piping Carries Rad Waste", GetFieldDataInt("VRadWaste", ref RRSus),
                                       "RadWaste", "VRadWaste");
                        _items.AddItem("Line Section is Governed by NSIAC", GetFieldDataInt("VNSIAC", ref RRSus),
                                       "NSIAC", "VNSIAC");
                        _items.AddItem("Safety Related Service Water", GetFieldDataInt("VSafetyRelated", ref RRSus),
                                       "SafetyRelated", "VSafetyRelated");
                        _items.AddItem("LCO <72 Hours", 1, "LCO", "VLCOLT72");
                        _items.AddItem("LCO 72 Hours", 2, "LCO", "VLCO72");
                        _items.AddItem("LCO 7 Day", 3, "LCO", "VLCO7");
                        _items.AddItem("LCO 14 Day", 4, "LCO", "VLCO14");
                        _items.AddItem("LCO 30 Day", 5, "LCO", "VLCO30");

                        _items.AddItem("Consequence Eng Judgment", 0, "Consequence Eng Judgment Value",
                                       "Consequence Eng Judgment Value");
                    }
                    connection.Close();
                }
                catch (Exception ex)
                {
                    if (ex.Message.Length > 0)
                        return;
                }
            }
        }
    }



    // ------------------------------------------------------------
    // ------------------------------------------------------------
    // ------------------------------------------------------------
    // BPSoilRiskRankingValues
    // ------------------------------------------------------------
    // ------------------------------------------------------------
    // ------------------------------------------------------------
    public class BPSoilRiskRankingValues
    {
        private BPRiskRankingValues _items;
        private long _lineid = 0;
        private long _beginpt = 0;
        private long _endpt = 0;
        public string _connectionstring = "";

        public BPRiskRankingValues RrValues
        {
            get { return _items; }
            set { }
        }

        // Constructor
        public BPSoilRiskRankingValues(long LineID, long BeginPoint, long EndPoint, string cs)
        {
            if ((LineID > 0) && (BeginPoint >= 0) && (EndPoint >= 0) && (cs.Length > 0))
            {
                _connectionstring = cs;
                _lineid = LineID;
                _beginpt = BeginPoint;
                _endpt = EndPoint;
                _items = new BPRiskRankingValues();
                LoadTheStructures();
            }
        }


        //-----------------------------------------------------------------------------------
        // See if the field exists
        //-----------------------------------------------------------------------------------
        public bool ColumnExists(IDataReader reader, string columnName)
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (string.Compare(columnName, reader.GetName(i)) == 0)
                {
                    return true;
                }
            }

            return false;
        }


        //-----------------------------------------------------------------------------------
        // Get BOOLEAN Field Data from the Reader
        //-----------------------------------------------------------------------------------
        private bool GetFieldDataBool(string DataFieldName, ref OleDbDataReader rdr, bool defvalue)
        {
            bool retval = defvalue;
            try
            {
                int idx = rdr.GetOrdinal(DataFieldName);
                if (idx >= 0)
                    retval = (string.Compare(rdr[idx].ToString(), "true", true) == 0);
            }
            finally
            {
            }
            return retval;
        }

        //-----------------------------------------------------------------------------------
        // Get STRING Field Data from the Reader
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
        // Get INT Field Data from the Reader
        //-----------------------------------------------------------------------------------
        private int GetFieldDataInt(string DataFieldName, ref OleDbDataReader rdr)
        {
            int retval = -1;
            try
            {
                int idx = rdr.GetOrdinal(DataFieldName);
                if (idx >= 0)
                {
                    retval = Convert.ToInt32(rdr[idx].ToString());
                }
            }
            finally
            {
            }
            return retval;
        }


        // Load the Values
        private void LoadTheStructures()
        {
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                string queryString = "Select * from [Dbf Setup - RiskRanking Con]";
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    double idx = 0;
                    OleDbDataReader RRSus = command.ExecuteReader();
                    if (RRSus.Read())
                    {
                        queryString = "Select * from [Dbf Consequence] WHERE LineID= " + _lineid +
                                      " AND ([Begin Point] <= " + _beginpt + " AND [End Point] >= " + _endpt + ")";
                        command = new OleDbCommand(queryString, connection);
                        OleDbDataReader LSCon = command.ExecuteReader();
                        if (LSCon.Read())
                        {
                            _items.Clear();
                            if (GetFieldDataBool("Licensing", ref LSCon, false))
                                _items.AddItem("Piping Has Specific Licensing Commitments",
                                               GetFieldDataInt("VLicensing", ref RRSus), "Licensing", "VLicensing");
                            if (GetFieldDataBool("Repair1Million", ref LSCon, false))
                                _items.AddItem("Repair > 1 Million Dollars",
                                               GetFieldDataInt("VRepair1Million", ref RRSus), "Repair1Million",
                                               "VRepair1Million");
                            if (GetFieldDataBool("Radiological", ref LSCon, false))
                                _items.AddItem("Piping Contains Radiological Fluid",
                                               GetFieldDataInt("VRadiological", ref RRSus), "Radiological",
                                               "VRadiological");
                            if (GetFieldDataBool("EPA", ref LSCon, false))
                                _items.AddItem("Piping Contains EPA Sensitive Fluid", GetFieldDataInt("VEPA", ref RRSus),
                                               "EPA", "VEPA");
                            if (GetFieldDataBool("UnderBldg2", ref LSCon, false))
                                _items.AddItem("Undermine Building or Supports of Other Systems",
                                               GetFieldDataInt("VUnderBuilding2", ref RRSus), "UnderBldg2",
                                               "VUnderBuilding2");
                            if (GetFieldDataBool("tritium", ref LSCon, false))
                                _items.AddItem("Process Pipe Carries Tritium", GetFieldDataInt("Vtritium", ref RRSus),
                                               "tritium", "Vtritium");
                            if (GetFieldDataBool("RadWaste", ref LSCon, false))
                                _items.AddItem("Process Piping Carries Rad Waste",
                                               GetFieldDataInt("VRadWaste", ref RRSus), "RadWaste", "VRadWaste");
                            if (GetFieldDataBool("NSIAC", ref LSCon, false))
                                _items.AddItem("Line Section is Governed by NSIAC", GetFieldDataInt("VNSIAC", ref RRSus),
                                               "NSIAC", "VNSIAC");
                            if (GetFieldDataBool("SafetyRelated", ref LSCon, false))
                                _items.AddItem("Safety Related Service Water",
                                               GetFieldDataInt("VSafetyRelated", ref RRSus), "SafetyRelated",
                                               "VSafetyRelated");

                            // LCO
                            int n = GetFieldDataInt("LCO", ref LSCon);
                            switch (n)
                            {
                                case 1:
                                    _items.AddItem("LCO <72 Hours", GetFieldDataInt("VLCOLT72", ref RRSus), "LCO",
                                                   "VLCOLT72");
                                    break;
                                case 2:
                                    _items.AddItem("LCO 72 Hours", GetFieldDataInt("VLCO72", ref RRSus), "LCO", "VLCO72");
                                    break;
                                case 3:
                                    _items.AddItem("LCO 7 Day", GetFieldDataInt("VLCO7", ref RRSus), "LCO", "VLCO7");
                                    break;
                                case 4:
                                    _items.AddItem("LCO 14 Day", GetFieldDataInt("VLCO14", ref RRSus), "LCO", "VLCO14");
                                    break;
                                case 5:
                                    _items.AddItem("LCO 30 Day", GetFieldDataInt("VLCO30", ref RRSus), "LCO", "VLCO30");
                                    break;
                            }
                            // Engineering Judgement
                            n = GetFieldDataInt("Consequence Eng Judgment Value", ref LSCon);
                            if (n > 0)
                                _items.AddItem("Consequence Eng Judgment",
                                               GetFieldDataInt("Consequence Eng Judgment Value", ref LSCon),
                                               "Consequence Eng Judgment Value", "Consequence Eng Judgment Value");
                            // Done
                            LSCon.Close();
                        }
                    }
                    RRSus.Close();
                }
                catch (Exception ex)
                {
                    if (ex.Message.Length > 0)
                        return;
                }
                connection.Close();
            }
        }
    }

}










//    MSql = "Select * from [Temp RR Soil]"
   //Set RstValue = CurrentDb.OpenRecordset(MSql, dbOpenDynaset)
   
   //If Not IsNull(pH) Then
   //   MSql = "Select * from [Dbf AnalysisData] Where [DataStart] <= " & pH & " and [DataEnd] >= " & pH & " and [DataType] = 'pH'"
   //   Set RR = CurrentDb.OpenRecordset(MSql, dbOpenSnapshot)
   //   If Not RR.EOF Then
   //      RstValue.AddNew
   //      RstValue("DataType") = RR("DataType")
   //      RstValue("DataDescr") = RR("DataDescr")
   //      RstValue("RiskPoints") = RR("RiskPoints")
   //      RstValue.Update
   //   End If
   //   RR.Close
   //   Set RR = Nothing
   //   DoEvents
   //End If
   //If Not IsNull(Cl) Then
   //   MSql = "Select * from [Dbf AnalysisData] Where [DataStart] <= " & Cl & " and [DataEnd] >= " & Cl & " and [DataType] = 'Chloride Content'"
   //   Set RR = CurrentDb.OpenRecordset(MSql, dbOpenSnapshot)
   //   If Not RR.EOF Then
   //      RstValue.AddNew
   //      RstValue("DataType") = RR("DataType")
   //      RstValue("DataDescr") = RR("DataDescr")
   //      RstValue("RiskPoints") = RR("RiskPoints")
   //      RstValue.Update
   //   End If
   //   RR.Close
   //   Set RR = Nothing
   //   DoEvents
   //End If
   //If Not IsNull(Resistivity) Then
   //   MSql = "Select * from [Dbf AnalysisData] Where [DataStart] <= " & Resistivity & " and [DataEnd] >= " & Resistivity & " and [DataType] = 'Soil Resistivity'"
   //   Set RR = CurrentDb.OpenRecordset(MSql, dbOpenSnapshot)
   //   If Not RR.EOF Then
   //      RstValue.AddNew
   //      RstValue("DataType") = RR("DataType")
   //      RstValue("DataDescr") = RR("DataDescr")
   //      RstValue("RiskPoints") = RR("RiskPoints")
   //      RstValue.Update
   //   End If
   //   RR.Close
   //   Set RR = Nothing
   //   DoEvents
   //End If
   //If Not IsNull(Native_Potential) Then
   //   MSql = "Select * from [Dbf AnalysisData] Where [DataStart] <= " & Native_Potential & " and [DataEnd] >= " & Native_Potential & " and [DataType] = 'Redox Potential'"
   //   Set RR = CurrentDb.OpenRecordset(MSql, dbOpenSnapshot)
   //   If Not RR.EOF Then
   //      RstValue.AddNew
   //      RstValue("DataType") = RR("DataType")
   //      RstValue("DataDescr") = RR("DataDescr")
   //      RstValue("RiskPoints") = RR("RiskPoints")
   //      RstValue.Update
   //   End If
   //   RR.Close
   //   Set RR = Nothing
   //   DoEvents
   //End If

   //If Not IsNull(Soil_Type) Then
   //   MSql = "Select * from [Dbf AnalysisData] Where [DataDescr] = '" & Soil_Type & "' and [DataType] = 'Soil Description'"
   //   Set RR = CurrentDb.OpenRecordset(MSql, dbOpenSnapshot)
   //   If Not RR.EOF Then
   //      RstValue.AddNew
   //      RstValue("DataType") = RR("DataType")
   //      RstValue("DataDescr") = RR("DataDescr")
   //      RstValue("RiskPoints") = RR("RiskPoints")
   //      RstValue.Update
   //   End If
   //   RR.Close
   //   Set RR = Nothing
   //   DoEvents
   //End If

