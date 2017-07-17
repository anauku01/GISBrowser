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
   
    public partial class BPLineDetails : Form
    {

        private bool _allow_open_line_sections = false;
        private string _connectionstring = "";
        private int _lineid = 0;

        //-----------------------------------------------------------------------------------
        // Contructor
        //-----------------------------------------------------------------------------------
        public BPLineDetails(int Line_ID, string connectionstring, bool allow_open_line_sections)
        {
            InitializeComponent();
            _allow_open_line_sections = allow_open_line_sections;
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
                    "SELECT [Dbf Lines].LineID, [Dbf Lines].Line, [Dbf Lines].Unit, [Dbf Lines].LineNumber, [Dbf Lines].LineDescription, [Dbf Lines].LineClass, [Dbf Lines].LineGroup, [Dbf Lines].LinePhase, [Dbf Lines].LineSafetyGrade, [Dbf Lines].LineSize, [Dbf Lines].BoreClass, [Dbf Lines].Category, [Dbf Lines].Criteria, [Dbf Lines].BasisComments, [Dbf Lines].FailureConsequence, [Dbf Lines].[Level of Susceptability], [Dbf Lines].Temperature, [Dbf Lines].[Line Content], [Dbf Lines].[Operating Time], [Dbf Lines].EvalComments, [Dbf Lines].SNMComments, [Dbf Lines].References, [Dbf Lines].PlantExperience, [Dbf Lines].PlantExperienceDescription, [Dbf Lines].IndustryExperience, [Dbf Lines].IndustryExperienceDescription, [Dbf Lines].CheckWorksLineName, [Dbf Lines].CheckWorksLineNumber, [Dbf Lines].[Failure Affects Safe Shutdown or CDF], [Dbf Lines].[Radiological Content], [Dbf Lines].Inservice, [Dbf Lines].[Tank Type], [Dbf Lines].[Begin Point], [Dbf Lines].[End Point], [Dbf Lines].CompanyID, [Dbf Lines].GUID, [Dbf Lines].OtherSpec, [Dbf Lines].Operation, [Dbf Lines].Velocity, [Dbf Lines].NSIAC, [Dbf Lines].LicenseRenewalCommitment, [Dbf Lines].CommitmentNumber, [Dbf Lines].LRCRemarks, [Dbf Lines].IsGWPILine, [Dbf Lines].[Total Length]";
                queryString = queryString + " FROM [Dbf Lines] WHERE [Dbf Lines].LineID =" + Line_ID.ToString() + ";";
                command = new OleDbCommand(queryString, connection);
                try
                {
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        // Fill in controls from the Lines Table
                        tbLine.Text = GetFieldDataStr("Line", ref reader);
                        tbLineNumber.Text = GetFieldDataStr("LineNumber", ref reader);
                        tbBeginPoint.Text = GetFieldDataStr("Begin Point", ref reader);
                        tbEndPoint.Text = GetFieldDataStr("End Point", ref reader);
                        tbTotalLength.Text = GetFieldDataStr("Total Length", ref reader);
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
                        tbTankType.Text = GetFieldDataStr("Tank Type", ref reader);
                        tbCriteria.Text = GetFieldDataStr("Criteria", ref reader);

                        tbCommitmentNumber.Text = GetFieldDataStr("CommitmentNumber", ref reader);
                        tbLRCComments.Text = GetFieldDataStr("LRCRemarks", ref reader);
                        tbBasisComments.Text = GetFieldDataStr("BasisComments", ref reader);
                        tbEvalComments.Text = GetFieldDataStr("EvalComments", ref reader);
                        tbPlantExperienceDesc.Text = GetFieldDataStr("PlantExperienceDescription", ref reader);
                        tbIndustryExperienceDesc.Text = GetFieldDataStr("IndustryExperienceDescription", ref reader);

                        contentlookupid = Convert.ToInt32(GetFieldDataStr("Line Content", ref reader));
                        lineclasslookupid = Convert.ToInt32(GetFieldDataStr("LineClass", ref reader));

                        // Boolean Fields
                        cbLicenseRenewalCommitment.Checked =
                            (Convert.ToBoolean(GetFieldDataStr("LicenseRenewalCommitment", ref reader)));
                        cbRadiologicalContent.Checked =
                            (Convert.ToBoolean(GetFieldDataStr("Radiological Content", ref reader)));
                        cbFailureAffectsShutdownOrCDF.Checked =
                            (Convert.ToBoolean(GetFieldDataStr("Failure Affects Safe Shutdown or CDF", ref reader)));
                        cbInService.Checked =
                            (Convert.ToBoolean(GetFieldDataStr("InService", ref reader)));
                        cbGWPILine.Checked =
                            (Convert.ToBoolean(GetFieldDataStr("IsGWPILine", ref reader)));
                        cbPlantExperience.Checked =
                            (Convert.ToBoolean(GetFieldDataStr("PlantExperience", ref reader)));
                        cbIndustryExperience.Checked =
                            (Convert.ToBoolean(GetFieldDataStr("IndustryExperience", ref reader)));
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


            // Fill in controls from Other Tables / Lookups
            //...............
            // Line Content
            if (contentlookupid > 0)
            {
                connection = new OleDbConnection(_connectionstring);
                if (connection != null)
                {
                    connection.Open();
                    queryString = "SELECT [Dbf Setup -  Data].SetupID, [Dbf Setup -  Data].ValueData, [Dbf Setup -  Data].Type, [Dbf Setup -  Data].Description, [Dbf Setup -  Data].GUID, [Dbf Setup -  Data].GenericID FROM [Dbf Setup -  Data] WHERE (([Dbf Setup -  Data].SetupID)=" + contentlookupid + ") AND (([Dbf Setup -  Data].Type)=\"content\")";
                    command = new OleDbCommand(queryString, connection);
                    try
                    {
                        reader = command.ExecuteReader();
                        if (reader.Read())
                        tbLineContent.Text = GetFieldDataStr("ValueData", ref reader);
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                    }
                    connection.Close();
                } // if connection!=null
            } // if contentlookupid >0


            //...............
            // Line Class
            //...............
            if (lineclasslookupid > 0)
            {
                queryString = "SELECT [Dbf Setup -  Data].SetupID, [Dbf Setup -  Data].ValueData, [Dbf Setup -  Data].Type, [Dbf Setup -  Data].Description, [Dbf Setup -  Data].GUID, [Dbf Setup -  Data].GenericID FROM [Dbf Setup -  Data] WHERE (([Dbf Setup -  Data].SetupID)=" + lineclasslookupid + ") AND (([Dbf Setup -  Data].Type)=\"LineCodeClass\")";
                command = new OleDbCommand(queryString, connection);
                try
                {
                    reader = command.ExecuteReader();
                    if (reader.Read())
                        tbLineClass.Text = GetFieldDataStr("ValueData", ref reader);
                    reader.Close();
                }
                catch (Exception ex)
                {
                }
                connection.Close();
            } // if lineclasslookupid>0          


            //...............
            // Line Section GRID Information
            //...............
            queryString = "SELECT [Dbf Line Section].[Line Section], [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point], [Dbf Line Section].[Overall Ranking], [Dbf Line Section].SusRanking, [Dbf Line Section].ConRanking, [Dbf Line Section].EngSusceptability, [Dbf Line Section].EngConsequence, [Dbf Line Section].EngJudgement, [Dbf Line Section].[Zone ID] FROM [Dbf Line Section] WHERE ((([Dbf Line Section].LineID)=" + Line_ID.ToString() + "))";
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
            if (_allow_open_line_sections)
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
