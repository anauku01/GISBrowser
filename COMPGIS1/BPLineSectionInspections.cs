using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace COMPGIS1
{
    public partial class BPLineSectionInspections : Form
    {

        private string _connectionstring = "";
        private int _zoneid = 0;

        public BPLineSectionInspections(int Zone_ID, string connectionstring)
        {
            InitializeComponent();
            _connectionstring = connectionstring;
            if (_connectionstring.Length > 0)
            {
                _zoneid = Zone_ID;
                LoadFormData(_zoneid);
                this.Text = "Line Section Inspections - [" + tbLineSection.Text + "]";
            }
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

        //-----------------------------------------------------------------------------------
        // Load the Form Data
        //-----------------------------------------------------------------------------------
        public int LoadFormData(int Zone_ID)
        {
            int beginpoint = -1;
            int endpoint = -1;
            int lineid = -1;
            string queryString = "";
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);

            if (connection != null)
                try
                {
                    connection.Open();
                    string ZoneStr = Zone_ID.ToString();
                    queryString =
                        "SELECT [Dbf Line Section].[Zone ID], [Dbf Line Section].[Line Section], [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point], [Dbf Line Section].Unit, [Dbf Line Section].[Line Section Ref #], [Dbf Line Section].[Line Section Type], [Dbf Line Section].[Total Length], [Dbf Line Section].[Line Section], [Dbf Line Section].[Line Section Description], [Dbf Line Section].Node, [Dbf Line Section].Comments, [Dbf Line Section].[Date Install], [Dbf Line Section].Line_Content, [Dbf Line Section].SusRanking, [Dbf Line Section].ConRanking, [Dbf Line Section].EngSusceptability, [Dbf Line Section].EngConsequence, [Dbf Line Section].EngJudgement, [Dbf Line Section].[Overall Ranking], [Dbf Line Section].[Contamination Level], [Dbf Line Section].Contaminated, [Dbf Line Section].[Contamination Type], [Dbf Line Section].[Dose Rate], [Dbf Line Section].Dose, [Dbf Line Section].[Zone Configuration], [Dbf Line Section].Coordinates, [Dbf Line Section].Area, [Dbf Line Section].[Area Description], [Dbf Line Section].[ASME Class], [Dbf Line Section].Building, [Dbf Line Section].[Building Zone], [Dbf Line Section].Elevation, [Dbf Line Section].Location, [Dbf Line Section].[Location Comments], [Dbf Line Section].Plant, [Dbf Line Section].[Safety Related], [Dbf Line Section].Loop, [Dbf Line Section].Room, [Dbf Line Section].[Replacement Comments], [Dbf Line Section].[Scaffolding Req], [Dbf Line Section].[Scaffolding Type], [Dbf Line Section].[Scaffolding Date], [Dbf Line Section].[Insulation Removal], [Dbf Line Section].[Insulation Removal Type], [Dbf Line Section].[Insulation Date], [Dbf Line Section].Preparation, [Dbf Line Section].[Preparation Type], [Dbf Line Section].[Preparation Date], [Dbf Line Section].[Grid Required], [Dbf Line Section].[Grid Type], [Dbf Line Section].GridTemplateSN, [Dbf Line Section].AxialGridSize, [Dbf Line Section].RadialGridSize, [Dbf Line Section].[Exam Instructions], [Dbf Line Section].[Accessible Online], [Dbf Line Section].[Inspection Priority], [Dbf Line Section].[Inspection LocationID], [Dbf Line Section].[Initial Inspection Number], [Dbf Line Section].[Installation Spec], [Dbf Line Section].[Inspection Interval], [Dbf Line Section].[Replacement WorkOrder], [Dbf Line Section].PrevComponentID, [Dbf Line Section].[Reinspection Interval], [Dbf Line Section].ReinstallationOutage, [Dbf Line Section].NextScheduledInspection, [Dbf Line Section].DefaultInspectionMeth, [Dbf Line Section].[Repair-Replace], [Dbf Line Section].[Repair-Replace Date1], [Dbf Line Section].[Repair-Replace WO1], [Dbf Line Section].[Repair-Replace Date2], [Dbf Line Section].[Repair-Replace WO2], [Dbf Line Section].[Repair-Replace Date3], [Dbf Line Section].[Repair-Replace WO3], [Dbf Line Section].Outage1, [Dbf Line Section].Outage2, [Dbf Line Section].Outage3, [Dbf Line Section].[Date Install USExt], [Dbf Line Section].[ReinstallationOutage USExt], [Dbf Line Section].[Date Install MainDS], [Dbf Line Section].[ReinstallationOutage MainDS], [Dbf Line Section].[Date Install DSExt], [Dbf Line Section].[ReinstallationOutage DSExt], [Dbf Line Section].[Date Install Branch], [Dbf Line Section].[ReinstallationOutage Branch], [Dbf Line Section].[Date Install BranchExt], [Dbf Line Section].[ReinstallationOutage BranchExt], [Dbf Line Section].[Line Section Status], [Dbf Line Section].Insulated, [Dbf Line Section].[Corrective Actions], [Dbf Line Section].[Preventative Actions], [Dbf Line Section].[Priority Description], [Dbf Line Section].Review, [Dbf Line Section].[Reviewed By], [Dbf Line Section].[Reviewed Date], [Dbf Line Section].[Pressure Class], [Dbf Line Section].Material_Type, [Dbf Line Section].Schedule_Type, [Dbf Line Section].Potential, [Dbf Line Section].CompanyID, [Dbf Line Section].SubUnitID, [Dbf Line Section].SubSystemID, [Dbf Line Section].CompanyName, [Dbf Line Section].SubSystemName, [Dbf Line Section].LineName, [Dbf Line Section].LineID " +
                        "FROM [Dbf Line Section] WHERE ((([Dbf Line Section].[Zone ID])=" + ZoneStr + "))";
                    command = new OleDbCommand(queryString, connection);
                    try
                    {
                        reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            // Fill in header controls from the Line Section Table
                            tbLineSection.Text = GetFieldDataStr("Line Section", ref reader);
                            tbBeginPoint.Text = GetFieldDataStr("Begin Point", ref reader);
                            tbEndPoint.Text = GetFieldDataStr("End Point", ref reader);
                            beginpoint = Convert.ToInt32(tbBeginPoint.Text);
                            endpoint = Convert.ToInt32(tbEndPoint.Text);
                            tbUnit.Text = GetFieldDataStr("Unit", ref reader);
                            // Lookup Fields
                            lineid = Convert.ToInt32(GetFieldDataStr("LineID", ref reader));
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        connection.Close();
                        return -1;
                    }
                    //...............
                    // General Inspections GRID Information
                    //...............
                    queryString = "SELECT [Dbf Line Section].[Zone ID], [Dbf Inspection Range].InspectionID, [Dbf Inspection Range].LineID, [Dbf Inspection Range].[Begin Point], " +
                                  "[Dbf Inspection Range].[End Point], [Dbf Inspection Range].Method, [Dbf Setup -  Data].ValueData AS [Insp Method], [Dbf Inspection Range].[Inspection Start Date], " +
                                  "[Dbf Inspection Range].[Inspection End Date], [Dbf Inspection Range].[Re-Inspection Date], [Dbf Inspection Range].Recommendations, [Dbf Inspection Range].Remarks, " +
                                  "[Dbf Inspection Range].GUID, [Dbf Inspection Range].Past_Inspections, [Dbf Inspection Range].Date_Past_Inspection, [Dbf Inspection Range].InspectionReportNo, " +
                                  "[Dbf Inspection Range].WallLossPerCent, [Dbf Inspection Range].NSI, [Dbf Inspection Range].Opportunistic, [Dbf Inspection Range].Directed, " +
                                  "[Dbf Inspection Range].InspectionReason FROM ([Dbf Line Section] LEFT JOIN [Dbf Inspection Range] ON [Dbf Line Section].LineID = [Dbf Inspection Range].LineID) " +
                                  "LEFT JOIN [Dbf Setup -  Data] ON [Dbf Inspection Range].Method = [Dbf Setup -  Data].SetupID " +
                                  "WHERE ((([Dbf Line Section].[Zone ID])=" + _zoneid + ") AND (([Dbf Inspection Range].[Begin Point])<=[Dbf Line Section]![Begin Point]) AND (([Dbf Inspection Range].[End Point])>=[Dbf Line Section]![End Point]));";


                    OleDbDataAdapter dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                    DataTable dTable = new DataTable();
                    //fill the DataTable
                    dAdapter.Fill(dTable);
                    DataGridView dgView = new DataGridView();
                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = dTable;
                    gridGeneralInspections.DataSource = bSource;
                    gridGeneralInspections.Columns["Zone ID"].Visible = false;
                    gridGeneralInspections.Columns["GUID"].Visible = false;
                    gridGeneralInspections.Columns["LineID"].Visible = false;
                    gridGeneralInspections.Columns["InspectionID"].Visible = false;
                    gridGeneralInspections.Columns["Method"].Visible = false;
                    gridGeneralInspections.Columns["Insp Method"].DisplayIndex = 0;


                    //...............
                    // Evaluations GRID Information
                    //...............
                    queryString = "SELECT [Dbf Inspections].[Zone ID], [Dbf Inspections].ID, [Dbf Line Section].[Line Section], [Dbf Inspections].Method, [Dbf Inspections].[WS Name], [Dbf Inspections].[Report No] " +
                                  "FROM [Dbf Inspections] INNER JOIN [Dbf Line Section] ON [Dbf Inspections].[Zone ID] = [Dbf Line Section].[Zone ID] " +
                                  "WHERE ((([Dbf Inspections].[Zone ID])=" + _zoneid.ToString() + "));";

                    dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                    dTable = new DataTable();
                    //fill the DataTable
                    dAdapter.Fill(dTable);
                    dgView = new DataGridView();
                    bSource = new BindingSource();
                    bSource.DataSource = dTable;
                    gridEvaluations.DataSource = bSource;
                    gridEvaluations.Columns["ID"].Visible = false;
                    gridEvaluations.Columns["Zone ID"].Visible = false;
                    gridEvaluations.Columns["Line Section"].Visible = false;
                    gridEvaluations.Columns["Method"].DisplayIndex = 0;
                    gridEvaluations.Columns["Report No"].DisplayIndex = 1;

                } // try...
                finally
                {
//                    connection.Close();
                }
            return 0;
        } // LoadFormData


        //-----------------------------------------------------------------------------------
        // Open the Evaluation Details Form
        //-----------------------------------------------------------------------------------
        private void OpenEvaluationForm(int InspectionID)
        {
            EvaluationDetails runform = new EvaluationDetails(InspectionID,_connectionstring);
            runform.ShowDialog();
            runform.Dispose();
        }

        //-----------------------------------------------------------------------------------
        // Open the Inspection Details Form
        //-----------------------------------------------------------------------------------
        private void gridGeneralInspections_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                int id = Convert.ToInt32(gridGeneralInspections.Rows[e.RowIndex].Cells[gridGeneralInspections.Columns["InspectionID"].Index].Value.ToString());
                if (id > 0)
                {
                    BPInspectionDetails runform = new BPInspectionDetails(id, _connectionstring);
                    runform.ShowDialog();
                }
            }
        }


        //-----------------------------------------------------------------------------------
        // Open the Evaluation Form
        //-----------------------------------------------------------------------------------
        private void gridEvaluations_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {            
            if (e.RowIndex >= 0)
            {                
                int id = Convert.ToInt32(gridEvaluations.Rows[e.RowIndex].Cells[gridEvaluations.Columns["ID"].Index].Value.ToString());
                string method = gridEvaluations.Rows[e.RowIndex].Cells[gridEvaluations.Columns["Method"].Index].Value.ToString();
                if (id > 0) OpenEvaluationForm(id);
            }
        } 


        //-----------------------------------------------------------------------------------
        // Close the Form
        //-----------------------------------------------------------------------------------
        private void btnClose_Click(object sender, EventArgs e)
        {
            // Close the Form
            this.Close();
        } // Close


    }
}
