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
    public partial class BEDetails : Form
    {

        private string _connectionstring = "";
        private string _boring_location_id = "";
        private bool _allow_open_chemical_results = false;
        private bool _allow_open_fieldlogdetails = true;
        private bool _allow_open_sieve_details = true;

        public BEDetails(string Boring_Location_ID, string connectionstring, bool allow_open_chemical_results)
        {
            InitializeComponent();
            if (Boring_Location_ID.Length > 0)
            {
                _allow_open_chemical_results = allow_open_chemical_results;
                _boring_location_id = Boring_Location_ID;
                _connectionstring = connectionstring;
                LoadFormData(_boring_location_id);
                this.Text = "Boring/Excavation Location Details - [" + _boring_location_id + "]";
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
        public int LoadFormData(string BoringLocationID)
        {
            string queryString = "";
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);

            if (connection != null)
            {
                connection.Open();
                try
                {

                    queryString =
                        "SELECT Soil_Boring_Locations.Boring_Location_ID, Soil_Boring_Locations.[Plant Northing], Soil_Boring_Locations.[Plant Easting], " +
                        "Soil_Boring_Locations.Location_Description, Soil_Boring_Locations.Groundwater_Level, Soil_Boring_Locations.IsExcavation, Soil_Boring_Locations.SampleDate, " +
                        "Soil_Boring_Locations.Notes FROM Soil_Boring_Locations WHERE (Soil_Boring_Locations.Boring_Location_ID=\"" + BoringLocationID + "\")";
                    command = new OleDbCommand(queryString, connection);
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        // Fill in controls from the Lines Table
                        tbBEID.Text = GetFieldDataStr("Boring_Location_ID", ref reader);
                        rbExcavation.Checked = (Convert.ToInt32(GetFieldDataStr("IsExcavation", ref reader)) != 0);
                        tbPlantEasting.Text = GetFieldDataStr("Plant Easting", ref reader);
                        tbPlantNorthing.Text = GetFieldDataStr("Plant Northing", ref reader);
                        tbGroundwaterLevel.Text = GetFieldDataStr("Groundwater_Level", ref reader);
                        tbDescription.Text = GetFieldDataStr("Location_Description", ref reader);
                        tbSampleDate.Text = GetFieldDataStr("SampleDate", ref reader);

                    }
                    reader.Close();

                }
                catch (Exception ex)
                {
                    connection.Close();
                    return -1;
                }


                //...............
                // Field Log GRID Information
                //...............
                queryString = "SELECT Soil_Field_Log.Field_Log_ID, Soil_Field_Log.SampledbDate, Soil_Field_Log.Boring_Location_ID, Soil_Field_Log.[Elevation (ft)], Soil_Field_Log.Sample_Depth_Top_FT_BGS, " +
                              "Soil_Field_Log.Sample_Depth_Btm_FT_BGS, Soil_Field_Log.Recovery_Ft, Soil_Field_Log.Boring_Description, Soil_Field_Log.Pipe_to_Soil_potential_V_SCE, Soil_Field_Log.Groundwater_Level, " +
                              "Soil_Field_Log.Ambient_Air_Temp FROM Soil_Field_Log WHERE ((Soil_Field_Log.Boring_Location_ID)=\"" + BoringLocationID + "\")";
                OleDbDataAdapter dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                DataTable dTable = new DataTable();
                //fill the DataTable
                dAdapter.Fill(dTable);
                DataGridView dgView = new DataGridView();
                BindingSource bSource = new BindingSource();
                bSource.DataSource = dTable;
                gridFieldLogs.DataSource = bSource;
                gridFieldLogs.Columns["Boring_Location_ID"].Visible = false;
                gridFieldLogs.Columns["Field_Log_ID"].Visible = false;
                gridFieldLogs.Columns["SampledbDate"].Visible = false;
               
                //...............
                // Corrosion Probe GRID Information
                //...............
                queryString = "SELECT Soil_Corrosion_Probes.Corrosion_Probe_ID, Soil_Corrosion_Probes.Boring_Location_ID, Soil_Corrosion_Probes.Depth, Soil_Corrosion_Probes.Serial_No, Soil_Corrosion_Probes.Make, " +
                              "Soil_Corrosion_Probes.Model, Soil_Corrosion_Probes.Installation_Method, Soil_Corrosion_Probes.ER_Coupon_Material, Soil_Corrosion_Probes.dbDate_Installed, Soil_Corrosion_Probes.Divisions_Baseline, " +
                              "Soil_Corrosion_Probes.Probe_Check_Baseline, Soil_Corrosion_Probes.Weather_Conditions FROM Soil_Corrosion_Probes WHERE ((Soil_Corrosion_Probes.Boring_Location_ID)=\"" + BoringLocationID + "\")";
                dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                dTable = new DataTable();
                //fill the DataTable
                dAdapter.Fill(dTable);
                dgView = new DataGridView();
                bSource = new BindingSource();
                bSource.DataSource = dTable;
                gridCP.DataSource = bSource;
                gridCP.Columns["Boring_Location_ID"].Visible = false;

                
                //...............
                // DCP GRID Information
                //...............
                queryString = "SELECT Soil_DCP_Results.DCP_Results_ID, Soil_DCP_Results.Boring_Location_ID, Soil_DCP_Results.Depth, Soil_DCP_Results.Run, Soil_DCP_Results.Drive_Distance_In, Soil_DCP_Results.DCP_Blows, " +
                              "Soil_DCP_Results.Estimated_Standard_N_Resistance_blows_ft, Soil_DCP_Results.Estimated_Cohesive_Soil_Consistency, Soil_DCP_Results.Total_Drive_Distance_in, Soil_DCP_Results.Notes FROM Soil_DCP_Results " +
                              "WHERE ((Soil_DCP_Results.Boring_Location_ID)=\"" + BoringLocationID + "\")";
                dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                dTable = new DataTable();
                //fill the DataTable
                dAdapter.Fill(dTable);
                dgView = new DataGridView();
                bSource = new BindingSource();
                bSource.DataSource = dTable;
                gridDCP.DataSource = bSource;
                gridDCP.Columns["Boring_Location_ID"].Visible = false;
                gridDCP.Columns["DCP_Results_ID"].Visible = false;
                

                //...............
                // Chemical Analysis Results GRID Information
                //...............
                queryString = "SELECT Soil_Chemical_Results.Soil_Chemical_Results_ID, Soil_Chemical_Results.Boring_Location_ID, Soil_Chemical_Results.Z, Soil_Chemical_Results.Sample_Depth_Top_FT_BGS, " +
                              "Soil_Chemical_Results.Sample_Depth_Btm_FT_BGS, Soil_Chemical_Results.CollectionDate, Soil_Chemical_Results.AnalysisDate, Soil_Chemical_Results.LabName, Soil_Chemical_Results.AnalysisMethod, " +
                              "Soil_Chemical_Results.Extractable_Calcium_mg_kg, Soil_Chemical_Results.Calcium_Lab_Result_Qualifier, Soil_Chemical_Results.Calcium_Reporting_Limit, Soil_Chemical_Results.Extractable_Magnesium_mg_kg, " +
                              "Soil_Chemical_Results.Magnesium_Lab_Result_Qualifier, Soil_Chemical_Results.Magnesium_Reporting_Limit, Soil_Chemical_Results.Extractable_Potassium_mg_kg, Soil_Chemical_Results.Potassium_Lab_Result_Qualifier, " +
                              "Soil_Chemical_Results.Potassium_Reporting_Limit, Soil_Chemical_Results.Extractable_Sodium_mg_kg, Soil_Chemical_Results.Sodium_Lab_Result_Qualifier, Soil_Chemical_Results.Sodium_Reporting_Limit, " +
                              "Soil_Chemical_Results.Extractable_Chloride_mg_kg, Soil_Chemical_Results.Chloride_Lab_Result_Qualifier, Soil_Chemical_Results.Chloride_Reporting_Limit, Soil_Chemical_Results.Extractable_Sulfate_mg_kg, " +
                              "Soil_Chemical_Results.Sulfate_Lab_Result_Qualifier, Soil_Chemical_Results.Sulfate_Reporting_Limit, Soil_Chemical_Results.Extractable_Sulfides_mg_kg, Soil_Chemical_Results.Sulfides_Lab_Result_Qualifier, " +
                              "Soil_Chemical_Results.Sulfides_Reporting_Limit, Soil_Chemical_Results.Sulfides_AWWA_Rating, Soil_Chemical_Results.pH_SU, Soil_Chemical_Results.pH_SU_Lab_Result_Qualifier, Soil_Chemical_Results.pH_SU_Reporting_Limit, " +
                              "Soil_Chemical_Results.pH_SU_AWWA_Rating, Soil_Chemical_Results.ORP_mV, Soil_Chemical_Results.ORP_mV_Lab_Result_Qualifier, Soil_Chemical_Results.ORP_mV_Reporting_Limit, Soil_Chemical_Results.ORP_mV_AWWA_Rating, " +
                              "Soil_Chemical_Results.Resistivity_Ohm_cm, Soil_Chemical_Results.Resistivity_Ohm_cm_Lab_Result_Qualifier, Soil_Chemical_Results.Resistivity_Ohm_cm_Reporting_Limit, Soil_Chemical_Results.Resistivity_Ohm_cm_AWWA_Rating, " +
                              "Soil_Chemical_Results.Resistivity_Ohm_cm_Classification, Soil_Chemical_Results.pcMoisture_wet_wt, Soil_Chemical_Results.pcMoisture_wet_wt_Lab_Result_Qualifier, Soil_Chemical_Results.pcMoisture_wet_wt_Reporting_Limit, " +
                              "Soil_Chemical_Results.pcMoisture_wet_wt_AWWA_Rating, Soil_Chemical_Results.pcSolids_wet_wt, Soil_Chemical_Results.pcSolids_wet_wt_Lab_Result_Qualifier, Soil_Chemical_Results.pcSolids_wet_wt_Reporting_Limit, " +
                              "Soil_Chemical_Results.pcSolids_wet_wt_AWWA_Rating, Soil_Chemical_Results.[Pipe-to-Soil_Potential_mV], Soil_Chemical_Results.PtS_Potential_mV_Lab_Result_Qualifier, Soil_Chemical_Results.PtS_Potential_mV_Reporting_Limit, " +
                              "Soil_Chemical_Results.PtS_Potential_mV_AWWA_Rating, Soil_Chemical_Results.AWWA_Point_Total, Soil_Chemical_Results.SRB_Found, Soil_Chemical_Results.SRB_Amount, Soil_Chemical_Results.PolarPlotImage " +
                              "FROM Soil_Chemical_Results WHERE ((Soil_Chemical_Results.Boring_Location_ID)=\"" + BoringLocationID + "\") ORDER BY Soil_Chemical_Results.Z";
                dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                dTable = new DataTable();
                //fill the DataTable
                dAdapter.Fill(dTable);
                dgView = new DataGridView();
                bSource = new BindingSource();
                bSource.DataSource = dTable;
                gridChemResults.DataSource = bSource;
                gridChemResults.Columns["Boring_Location_ID"].Visible = false;
                gridChemResults.Columns["Soil_Chemical_Results_ID"].Visible = false;


                //...............
                // Line Section GRID Information
                //...............
                queryString = "SELECT [Soil - Line Section].[Zone ID], [Soil - Line Section].Boring_Location_ID, [Dbf Line Section].[Line Section], [Soil - Line Section].[Report No], [Soil - Line Section].[Work Scope] FROM [Soil - Line Section] " +
                              "INNER JOIN [Dbf Line Section] ON [Soil - Line Section].[Zone ID] = [Dbf Line Section].[Zone ID] WHERE ((([Soil - Line Section].Boring_Location_ID)=\"" + BoringLocationID + "\")) " +
                              "ORDER BY [Dbf Line Section].[Line Section]";
                dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                dTable = new DataTable();
                //fill the DataTable
                dAdapter.Fill(dTable);
                dgView = new DataGridView();
                bSource = new BindingSource();
                bSource.DataSource = dTable;
                gridLineSections.DataSource = bSource;
                gridLineSections.Columns["Boring_Location_ID"].Visible = false;
                gridLineSections.Columns["Zone ID"].Visible = false;


                //...............
                // Sieve Analysis GRID Information
                //...............
                queryString = "SELECT Soil_Sieve_Analysis.Soil_Sieve_Analysis_ID, Soil_Sieve_Analysis.Boring_Location_ID, Soil_Sieve_Analysis.Z, Soil_Sieve_Analysis.Sample_Depth_Top_FT_BGS, Soil_Sieve_Analysis.Sample_Depth_Btm_FT_BGS, " +
                              "Soil_Sieve_Analysis.Sieve_3_inch, Soil_Sieve_Analysis.Sieve_2_inch, Soil_Sieve_Analysis.Sieve_1_1_2_inch, Soil_Sieve_Analysis.Sieve_3_4_inch, Soil_Sieve_Analysis.Sieve_3_8_inch, Soil_Sieve_Analysis.Sieve_Num_4, " +
                              "Soil_Sieve_Analysis.Sieve_Num_10, Soil_Sieve_Analysis.Sieve_Num_20, Soil_Sieve_Analysis.Sieve_Num_40, Soil_Sieve_Analysis.Sieve_Num_60, Soil_Sieve_Analysis.Sieve_Num_100, Soil_Sieve_Analysis.Sieve_075, " +
                              "Soil_Sieve_Analysis.Sieve_023, Soil_Sieve_Analysis.Sieve_0136, Soil_Sieve_Analysis.Sieve_0098, Soil_Sieve_Analysis.Sieve_007, Soil_Sieve_Analysis.Sieve_0035, Soil_Sieve_Analysis.Sieve_0015, " +
                              "Soil_Sieve_Analysis.Gravel_pc, Soil_Sieve_Analysis.Sand_Course_pc, Soil_Sieve_Analysis.SandMedium_pc, Soil_Sieve_Analysis.SandFine_pc, Soil_Sieve_Analysis.Silt_pc, Soil_Sieve_Analysis.Clay_colloids_pc, " +
                              "Soil_Sieve_Analysis.Sieve_1_inch, Soil_Sieve_Analysis.Sieve_036 FROM Soil_Sieve_Analysis WHERE (((Soil_Sieve_Analysis.Boring_Location_ID)=\"" +  BoringLocationID + "\")) ORDER BY Soil_Sieve_Analysis.Z";
                dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                dTable = new DataTable();
                //fill the DataTable
                dAdapter.Fill(dTable);
                dgView = new DataGridView();
                bSource = new BindingSource();
                bSource.DataSource = dTable;
                gridSieve.DataSource = bSource;
                gridSieve.Columns["Boring_Location_ID"].Visible = false;
                gridSieve.Columns["Soil_Sieve_Analysis_ID"].Visible = false;


                //connection.Close();
            } // if connection!=null

            // DONE
            return 0;
        } // LoadFormData


        //-----------------------------------------------------------------------------------
        // Opens the Line Section dialog displaying the Line Section selected in the grid
        //----------------------------------------------------------------------------------- 
        private void gridLineSections_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                int id = Convert.ToInt32(gridLineSections.Rows[e.RowIndex].Cells[gridLineSections.Columns["Zone ID"].Index].Value.ToString());
                if (id > 0)
                {
                    BPLineSectionDetails runform = new BPLineSectionDetails(id, _connectionstring, true);
                    runform.ShowDialog();
                }               
            }
        }


        //-----------------------------------------------------------------------------------
        // Opens the Corrosion Probe Details form from the CPGrid doubleclick
        //-----------------------------------------------------------------------------------
        private void gridCP_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                string id = gridCP.Rows[e.RowIndex].Cells[gridCP.Columns["Corrosion_Probe_ID"].Index].Value.ToString();
                if (id.Length > 0)
                {
                    BPCP_Details runform = new BPCP_Details(id, _connectionstring);
                    runform.ShowDialog();
                }
            }
        }


        //-----------------------------------------------------------------------------------
        // Load the Chemical Analysis Results form
        //-----------------------------------------------------------------------------------
        private void gridChemResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (_allow_open_chemical_results)
            {
                if (e.RowIndex >= 0)
                {
                    int id = Convert.ToInt32(gridChemResults.Rows[e.RowIndex].Cells[gridChemResults.Columns["Soil_Chemical_Results_ID"].Index].Value.ToString());
                    if (id > 0)
                    {
                        BPChemicalAnalysisDetails runform = new BPChemicalAnalysisDetails(id, _connectionstring);
                        runform.ShowDialog();
                    }
                }
            }
        }


        //-----------------------------------------------------------------------------------
        // Load the Field Log Details form
        //-----------------------------------------------------------------------------------

        private void gridFieldLogs_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (_allow_open_fieldlogdetails)
            {
                if (e.RowIndex >= 0)
                {
                    int id = Convert.ToInt32(gridFieldLogs.Rows[e.RowIndex].Cells[gridFieldLogs.Columns["Field_Log_ID"].Index].Value.ToString());
                    if (id > 0)
                    {
                        FieldLogDetails runform = new FieldLogDetails(id, _connectionstring);
                        runform.ShowDialog();
                    }
                }
            }
        }


        //-----------------------------------------------------------------------------------
        // Load the Sieve Analysis Details
        //-----------------------------------------------------------------------------------
        private void gridSieve_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (_allow_open_sieve_details)
            {
                if (e.RowIndex >= 0)
                {
                    int id = Convert.ToInt32(gridSieve.Rows[e.RowIndex].Cells[gridSieve.Columns["Soil_Sieve_Analysis_ID"].Index].Value.ToString());
                    if (id > 0)
                    {
                        BPSieveAnalysisDetails runform = new BPSieveAnalysisDetails(id, _connectionstring);
                        runform.ShowDialog();
                    }
                }
            }
        }

        //-----------------------------------------------------------------------------------
        // Close the Form
        //-----------------------------------------------------------------------------------
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }




    }
}
