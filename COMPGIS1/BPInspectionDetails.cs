using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace COMPGIS1
{
    public partial class BPInspectionDetails : Form
    {

        private string _connectionstring = "";
        private int _inspectionid = 0;

        //-----------------------------------------------------------------------------------
        // Constructor
        //-----------------------------------------------------------------------------------
        public BPInspectionDetails(int inspectionID, string connectionstring)
        {
            InitializeComponent();
            _connectionstring = connectionstring;
            if (_connectionstring.Length > 0)
            {
                _inspectionid = inspectionID;
                LoadFormData(_inspectionid);
            }
        }


        //-----------------------------------------------------------------------------------
        // Get string Field Data from the Reader
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
        // Get int Field Data from the Reader
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

        //-----------------------------------------------------------------------------------
        // Load the Form Data
        //-----------------------------------------------------------------------------------
        public int LoadFormData(int inspectionID)
        {
            string queryString = "";
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);

            if (connection != null)
                try
                {
                    connection.Open();
                    queryString =
                    "SELECT [Dbf Inspection Range].InspectionID, [Dbf Lines].Line, [Dbf Inspection Range].[Begin Point], [Dbf Inspection Range].[End Point], " +
                    "[Dbf Setup -  Data].ValueData AS [Insp Method], [Dbf Inspection Range].[Inspection Start Date], [Dbf Inspection Range].[Inspection End Date], " + 
                    "[Dbf Inspection Range].[Re-Inspection Date], [Dbf Inspection Range].Recommendations, [Dbf Inspection Range].Remarks, [Dbf Inspection Range].GUID, " + 
                    "[Dbf Inspection Range].Past_Inspections, [Dbf Inspection Range].Date_Past_Inspection, [Dbf Inspection Range].InspectionReportNo, " + 
                    "[Dbf Inspection Range].WallLossPerCent, [Dbf Inspection Range].NSI, [Dbf Inspection Range].Opportunistic, [Dbf Inspection Range].Directed, " + 
                    "[Dbf Inspection Range].InspectionReason FROM ([Dbf Inspection Range] LEFT JOIN [Dbf Setup -  Data] ON [Dbf Inspection Range].Method = [Dbf Setup -  Data].SetupID) " + 
                    "INNER JOIN [Dbf Lines] ON [Dbf Inspection Range].LineID = [Dbf Lines].LineID WHERE ((([Dbf Inspection Range].InspectionID)=" + inspectionID.ToString() + "))";

                    command = new OleDbCommand(queryString, connection);
                    try
                    {
                        reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            tbLineNumber.Text = GetFieldDataStr("Line", ref reader);
                            tbBeginPoint.Text = GetFieldDataStr("Begin Point", ref reader);
                            tbEndPoint.Text = GetFieldDataStr("End Point", ref reader);
                            tbMethod.Text = GetFieldDataStr("Insp Method", ref reader);
                            tbInspStartDate.Text = GetFieldDataStr("Inspection Start Date", ref reader);
                            tbInspEndDate.Text = GetFieldDataStr("Inspection End Date", ref reader);
                            tbReinspDate.Text = GetFieldDataStr("Re-Inspection Date", ref reader);
                            tbDatePastInsp.Text = GetFieldDataStr("Date_Past_Inspection", ref reader);
                            tbRecommendations.Text = GetFieldDataStr("Recommendations", ref reader);
                            tbRemarks.Text = GetFieldDataStr("Remarks", ref reader);
                            tbPastInsp.Text = GetFieldDataStr("Past_Inspections", ref reader);
                            tbInspRepNum.Text = GetFieldDataStr("InspectionReportNo", ref reader);
                            tbWallLossPercent.Text = GetFieldDataStr("WallLossPerCent", ref reader);
                            tbNSI.Text = GetFieldDataStr("NSI", ref reader);
                            cbOpportunistic.Checked = (Convert.ToBoolean(GetFieldDataStr("Opportunistic", ref reader)));
                            cbDirected.Checked = (Convert.ToBoolean(GetFieldDataStr("Directed", ref reader)));
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        connection.Close();
                        return -1;
                    }

                    //...............
                    // Fill in the Generic Grid
                    //...............
                    queryString =   "SELECT [Dbf Inspection Generic].InspectionID, [Dbf Inspection Generic].[Begin Point], [Dbf Inspection Generic].[End Point], " + 
                                    "[Dbf Setup -  Data].ValueData AS InspectionType, [Dbf Setup -  Data_1].ValueData AS Disposition, [Dbf Inspection Generic].[Inspection Date], " +
                                    "[Dbf Inspection Generic].Tmin, [Dbf Inspection Generic].[Estimated Corrosion Rate], [Dbf Inspection Generic].Remarks, [Dbf Inspection Generic].GUID " +
                                    "FROM ([Dbf Inspection Generic] LEFT JOIN [Dbf Setup -  Data] ON [Dbf Inspection Generic].[Inspection Type] = [Dbf Setup -  Data].SetupID) LEFT JOIN [Dbf Setup -  Data] AS [Dbf Setup -  Data_1] ON [Dbf Inspection Generic].Disposition = [Dbf Setup -  Data_1].SetupID " +
                                    "WHERE ((([Dbf Inspection Generic].InspectionID)=" + inspectionID + "))";
                    OleDbDataAdapter dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                    DataTable dTable = new DataTable();
                    //fill the DataTable
                    dAdapter.Fill(dTable);
                    DataGridView dgView = new DataGridView();
                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = dTable;
                    gridGeneric.DataSource = bSource;

                    //...............
                    // Fill in the Anomaly Grid
                    //...............
                    queryString = "SELECT [Dbf Inspection Anomaly].InspectionID, [Dbf Inspection Anomaly].[Begin Point], [Dbf Inspection Anomaly].[End Point], [Dbf Setup -  Data].ValueData AS [Deposit Type], " + 
                    "[Dbf Inspection Anomaly].[Detect Date], [Dbf Inspection Anomaly].Type, [Dbf Setup -  Data_1].ValueData AS [Deposit Level], [Dbf Inspection Anomaly].Length, [Dbf Inspection Anomaly].Width, " + 
                    "[Dbf Inspection Anomaly].Depth, [Dbf Setup -  Data_2].ValueData AS Orientation, [Dbf Inspection Anomaly].Remarks, [Dbf Inspection Anomaly].GUID FROM (([Dbf Inspection Anomaly] " + 
                    "LEFT JOIN [Dbf Setup -  Data] ON [Dbf Inspection Anomaly].[Deposit Type] = [Dbf Setup -  Data].SetupID) LEFT JOIN [Dbf Setup -  Data] AS [Dbf Setup -  Data_1] ON [Dbf Inspection Anomaly].[Deposit Level] = [Dbf Setup -  Data_1].SetupID) " + 
                    "INNER JOIN [Dbf Setup -  Data] AS [Dbf Setup -  Data_2] ON [Dbf Inspection Anomaly].Orientation = [Dbf Setup -  Data_2].SetupID " +
                    "WHERE ((([Dbf Inspection Anomaly].InspectionID)=" + inspectionID + "))";

                    dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                    dTable = new DataTable();
                    //fill the DataTable
                    dAdapter.Fill(dTable);
                    dgView = new DataGridView();
                    bSource = new BindingSource();
                    bSource.DataSource = dTable;
                    gridAnomaly.DataSource = bSource;
                    gridAnomaly.Columns["InspectionID"].Visible = false;


                    //...............
                    // Fill in the UTWall Grid
                    //...............
                    queryString = "SELECT [Dbf Inspection UTWall].InspectionID, [Dbf Inspection UTWall].[Begin Point], [Dbf Inspection UTWall].[End Point], [Dbf Inspection UTWall].[Inspection Date], " + 
                    "[Dbf Inspection UTWall].[Min Wall], [Dbf Inspection UTWall].[Avg Wall], [Dbf Setup -  Data].ValueData AS Severity, [Dbf Inspection UTWall].Unit, [Dbf Inspection UTWall].[Unit Serial Number], " + 
                    "[Dbf Inspection UTWall].Remarks, [Dbf Inspection UTWall].Inspector, [Dbf Inspection UTWall].GUID FROM [Dbf Inspection UTWall] INNER JOIN [Dbf Setup -  Data] ON [Dbf Inspection UTWall].Severity = [Dbf Setup -  Data].SetupID " + 
                    "WHERE ((([Dbf Inspection UTWall].InspectionID)=" + inspectionID + "))";
                    dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                    dTable = new DataTable();
                    //fill the DataTable
                    dAdapter.Fill(dTable);
                    dgView = new DataGridView();
                    bSource = new BindingSource();
                    bSource.DataSource = dTable;
                    gridUTWall.DataSource = bSource;
                    gridUTWall.Columns["InspectionID"].Visible = false;

                    //...............
                    // Fill in the PCCP Grid
                    //...............
                    queryString = "SELECT [Dbf Inspection PCCP].InspectionID, [Dbf Inspection PCCP].[Begin Point], [Dbf Inspection PCCP].[End Point], [Dbf Inspection PCCP].[Inspection Date], [Dbf Setup -  Data].ValueData AS PipeCondition, "+
                    "[Dbf Setup -  Data_1].ValueData AS [Joint Condition], [Dbf Inspection PCCP].[Broken Wires], [Dbf Inspection PCCP].[Wires Inspected], [Dbf Inspection PCCP].Classification, [Dbf Inspection PCCP].Remarks, [Dbf Inspection PCCP].GUID, "+ 
                    "[Dbf Inspection PCCP].MortarCondition FROM ([Dbf Inspection PCCP] LEFT JOIN [Dbf Setup -  Data] ON [Dbf Inspection PCCP].[Pipe Condition] = [Dbf Setup -  Data].SetupID) LEFT JOIN [Dbf Setup -  Data] AS [Dbf Setup -  Data_1] ON [Dbf Inspection PCCP].[Joint Condition] = [Dbf Setup -  Data_1].SetupID " +
                    "WHERE ((([Dbf Inspection PCCP].InspectionID)=" + inspectionID + "))";
                    dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                    dTable = new DataTable();
                    //fill the DataTable
                    dAdapter.Fill(dTable);
                    dgView = new DataGridView();
                    bSource = new BindingSource();
                    bSource.DataSource = dTable;
                    gridPCCP.DataSource = bSource;
                    gridPCCP.Columns["InspectionID"].Visible = false;

                    //...............
                    // Fill in the GWT Grid
                    //...............
                    queryString = "SELECT [Dbf Inspections Guided Wave].InspectionID, [Dbf Inspections Guided Wave].[Begin Point], [Dbf Inspections Guided Wave].[End Point], [Dbf Inspections Guided Wave].[Inspection Date], [Dbf Inspections Guided Wave].ResultsLevel, [Dbf Inspections Guided Wave].ShotReference, " +
                    "[Dbf Inspections Guided Wave].Feature, [Dbf Inspections Guided Wave].Location, [Dbf Inspections Guided Wave].ECLPercent, [Dbf Inspections Guided Wave].CircAffectedPercent, [Dbf Inspections Guided Wave].ApprWallLossPercent, "+
                    "[Dbf Inspections Guided Wave].CorrosionLength, [Dbf Inspections Guided Wave].ResultClass, [Dbf Inspections Guided Wave].Extent, [Dbf Inspections Guided Wave].OtherFeature, [Dbf Inspections Guided Wave].Remarks " +
                    "FROM [Dbf Inspections Guided Wave] WHERE ((([Dbf Inspections Guided Wave].InspectionID)=" + inspectionID + "))";
                    dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                    dTable = new DataTable();
                    //fill the DataTable
                    dAdapter.Fill(dTable);
                    dgView = new DataGridView();
                    bSource = new BindingSource();
                    bSource.DataSource = dTable;
                    gridGWT.DataSource = bSource;
                    gridGWT.Columns["InspectionID"].Visible = false;

                } // try...
                finally
                {
                    connection.Close();
                }
            return 0;
        } // LoadFormData



        //-----------------------------------------------------------------------------------
        // Close the form
        //-----------------------------------------------------------------------------------
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        } 







    }
}
