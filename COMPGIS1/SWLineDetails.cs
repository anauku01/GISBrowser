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
   
    public partial class SWLineDetails : Form
    {

        private bool _allow_open_flow_segments = false;
        private string _connectionstring = "";
        private int _lineid = 0;

        //-----------------------------------------------------------------------------------
        // Contructor
        //-----------------------------------------------------------------------------------
        public SWLineDetails(int Line_ID, string connectionstring, bool allow_open_flow_segments)
        {
            InitializeComponent();
            _allow_open_flow_segments = allow_open_flow_segments;
            _connectionstring = connectionstring;
            if (_connectionstring.Length > 0)
            {
                _lineid = Line_ID;
                LoadFormData(_lineid);
                this.Text = "Line Details - [" + tbLine.Text + "]";                
            }
        }

        //-----------------------------------------------------------------------------------
        // Get Field Data from the Reader
        //-----------------------------------------------------------------------------------
        private string GetFieldDataStr(string DataFieldName, ref OleDbDataReader rdr)
        {
            int idx = rdr.GetOrdinal(DataFieldName);
            if (idx >= 0)
                return rdr[idx].ToString();
            else
                return "";
        }


        //-----------------------------------------------------------------------------------
        // Load the Form Data
        //-----------------------------------------------------------------------------------
        public int LoadFormData(int Line_ID)
        {
            int contentlookupid = -1;
            int lineclasslookupid = -1;
            string queryString = "";
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);

            if (connection != null)
            {
                connection.Open();
                queryString =
                    "SELECT * FROM [Dbf Lines] WHERE (([Dbf Lines].LineID)=" + Line_ID.ToString() + ")";

                //queryString =
                //    "SELECT [Dbf Lines].LineID, [Dbf Lines].Line, [Dbf Lines].LineNumber, [Dbf Lines].LineDescription, [Dbf Lines].LineClass, [Dbf Lines].LineGroup, [Dbf Lines].LinePhase, [Dbf Lines].LineSafetyGrade, " +
                //    "[Dbf Lines].LineSize, [Dbf Lines].BoreClass, [Dbf Lines].Category, [Dbf Lines].Criteria, [Dbf Lines].BasisComments, [Dbf Lines].CheckWorksLineName, [Dbf Lines].CheckWorksLineNumber, [Dbf Lines].FailureConsequence, " +
                //    "[Dbf Lines].[Level of Susceptability], [Dbf Lines].Temperature, [Dbf Lines].[Steam Quality], [Dbf Lines].[Operating Time], [Dbf Lines].EvalComments, [Dbf Lines].SNMComments, [Dbf Lines].References, " +
                //    "[Dbf Lines].PlantExperience, [Dbf Lines].PlantExperienceDescription, [Dbf Lines].IndustryExperience, [Dbf Lines].IndustryExperienceDescription, [Dbf Lines].Comments, [Dbf Lines].Rev, [Dbf Lines].RefNo, " +
                //    "[Dbf Lines].[Drawing No], [Dbf Lines].SystemID, [Dbf Lines].Unit, [Dbf Lines].Usage, [Dbf Lines].BeginPoint, [Dbf Lines].EndPoint, [Dbf Lines].MapFeatureID " +
                //    "FROM [Dbf Lines] WHERE (([Dbf Lines].LineID)=" + Line_ID.ToString() + ")";

                //queryString =
                //    "SELECT [Dbf Lines].LineID, [Dbf Lines].Line, [Dbf Lines].LineNumber, [Dbf Lines].LineDescription, [Dbf Lines].LineClass, [Dbf Systems].System, [Dbf Lines].LineGroup, [Dbf Lines].LinePhase, " +
                //    "[Dbf Lines].LineSafetyGrade, [Dbf Lines].LineSize, [Dbf Lines].BoreClass, [Dbf Lines].Category, [Dbf Lines].Criteria, [Dbf Lines].BasisComments, [Dbf Lines].CheckWorksLineName, [Dbf Lines].CheckWorksLineNumber, " +
                //    "[Dbf Lines].FailureConsequence, [Dbf Lines].[Level of Susceptability], [Dbf Lines].Temperature, [Dbf Lines].[Steam Quality], [Dbf Lines].[Operating Time], [Dbf Lines].EvalComments, [Dbf Lines].SNMComments, " +
                //    "[Dbf Lines].References, [Dbf Lines].PlantExperience, [Dbf Lines].PlantExperienceDescription, [Dbf Lines].IndustryExperience, [Dbf Lines].IndustryExperienceDescription, [Dbf Lines].Comments, [Dbf Lines].Rev, " +
                //    "[Dbf Lines].RefNo, [Dbf Lines].[Drawing No], [Dbf Lines].Unit, [Dbf Lines].Usage, [Dbf Lines].BeginPoint, [Dbf Lines].EndPoint, [Dbf Lines].MapFeatureID " +
                //    "FROM [Dbf Lines] INNER JOIN [Dbf Systems] ON [Dbf Lines].SystemID = [Dbf Systems].SystemID WHERE ((([Dbf Lines].LineID)=" + Line_ID.ToString() + "));";
                    
                command = new OleDbCommand(queryString, connection);
                try
                {
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        // Fill in controls from the Lines Table
                        tbLine.Text = GetFieldDataStr("Line", ref reader);
                        tbLineNumber.Text = GetFieldDataStr("LineNumber", ref reader);
                        tbBeginPoint.Text = GetFieldDataStr("BeginPoint", ref reader);
                        tbEndPoint.Text = GetFieldDataStr("EndPoint", ref reader);
                        tbLineDescription.Text = GetFieldDataStr("LineDescription", ref reader);
                        tbLineGroup.Text = GetFieldDataStr("LineGroup", ref reader);
                        tbLineGrade.Text = GetFieldDataStr("LineSafetyGrade", ref reader);
                        tbUnit.Text = GetFieldDataStr("Unit", ref reader);
                        tbLineSize.Text = GetFieldDataStr("LineSize", ref reader);
                        tbTemperature.Text = GetFieldDataStr("Temperature", ref reader);
                        tbOperatingTime.Text = GetFieldDataStr("Operating Time", ref reader);
                        tbLinePhase.Text = GetFieldDataStr("LinePhase", ref reader);
                        tbLineCategory.Text = GetFieldDataStr("Category", ref reader);
                        tbReferences.Text = GetFieldDataStr("References", ref reader);
                        tbCriteria.Text = GetFieldDataStr("Criteria", ref reader);
                        tbComments.Text = GetFieldDataStr("Comments", ref reader);
                    }
                reader.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                    return -1;
                }
            connection.Close();
            } // if connection!=null

            //...............
            // Flow Segment GRID Information
            //...............
            queryString =
                "SELECT [Dbf Zone].[Zone ID], [Dbf Zone].[Flow Segment], [Dbf Zone].[Flow Segment Ref #], [Dbf Zone].[Flow Segment Type], [Dbf Zone].[Flow Segment Description], [Dbf Systems].System, [Dbf Lines].Line, [Dbf Zone].Node, [Dbf Zone].[Flow (gpm)], [Dbf Zone].[Flow Density], [Dbf Zone].[Nominal Pipe Size], [Dbf Zone].[Pipe OD], [Dbf Zone].[Pipe ID], [Dbf Zone].Area, [Dbf Zone].Velocity, [Dbf Zone].Comments, [Dbf Zone].Scheduled, [Dbf Zone].PID, " +
                "[Dbf Zone].Iso, [Dbf Zone].PRA, [Dbf Zone].LCO, [Dbf Zone].Accessablility, [Dbf Zone].LineSize, [Dbf Zone].Building, [Dbf Zone].Elevation, [Dbf Zone].Row, [Dbf Zone].Col, [Dbf Zone].SR, [Dbf Zone].Schedule, [Dbf Zone].Category, [Dbf Zone].ChemTF, [Dbf Zone].Online, [Dbf Zone].SRB, [Dbf Zone].TAB, [Dbf Zone].APB, [Dbf Zone].IRB, [Dbf Zone].Cl, [Dbf Zone].SO3, [Dbf Zone].HCO2, [Dbf Zone].CO2, [Dbf Zone].Comments2, [Dbf Zone].InspLocation, " +
                "[Dbf Zone].Drawing1, [Dbf Zone].Drawing2, [Dbf Zone].Unit, [Dbf Zone].[Tnom Main], [Dbf Zone].BeginPoint, [Dbf Zone].EndPoint, [Dbf Zone].MapFeatureID FROM ([Dbf Zone] INNER JOIN [Dbf Lines] ON [Dbf Zone].LineID = [Dbf Lines].LineID) INNER JOIN [Dbf Systems] ON [Dbf Lines].SystemID = [Dbf Systems].SystemID WHERE ((([Dbf Zone].[Zone ID])=" + Line_ID.ToString() + "))";            
            OleDbDataAdapter dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
            //OleDbCommandBuilder cBuilder = new OleDbCommandBuilder(dAdapter);
            DataTable dTable = new DataTable();
            //fill the DataTable
            dAdapter.Fill(dTable);
            //the DataGridView
            //DataGridView dgView = new DataGridView();
            BindingSource bSource = new BindingSource();
            bSource.DataSource = dTable;
            gridLineSections.DataSource = bSource;
            gridLineSections.Columns["Zone ID"].Visible = false;

        // DONE
        return 0;
        } // loadformdata 


        //-----------------------------------------------------------------------------------
        // User Double Clicks on Line Section grid - displays Line Section Form for selected Line Section
        //-----------------------------------------------------------------------------------
        private void gridLineSections_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (_allow_open_flow_segments)
            {
                if (e.RowIndex >= 0)
                {
                    int zoneid = Convert.ToInt32(gridLineSections.Rows[e.RowIndex].Cells[gridLineSections.Columns["Zone ID"].Index].Value.ToString());
                    if (zoneid > 0)
                    {                       
                        BPLineSectionDetails lsform = new BPLineSectionDetails(zoneid, _connectionstring, false);
                        lsform.ShowDialog();
                    }
                }
            }
        }


        //-----------------------------------------------------------------------------------
        // Close the Form
        //-----------------------------------------------------------------------------------
        private void btnClose_Click(object sender, EventArgs e)
        {
            // Close the Form
            this.Close();
        }



    }

}
