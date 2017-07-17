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
    //----------------------------------------------------------------------------------------------------
    // Constructor
    //----------------------------------------------------------------------------------------------------
    public partial class BPCP_Details : Form
    {

        private string _connectionstring = "";
        private string _corrosion_probe_id = "";

        public BPCP_Details(string Corrosion_Probe_ID, string connectionstring)
        {
            InitializeComponent();
            _corrosion_probe_id = Corrosion_Probe_ID;
            _connectionstring = connectionstring;
            LoadFormData();
            this.Text = "Corrosion Probe Details";            
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
        public bool LoadFormData()
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
                    queryString = "SELECT Soil_Corrosion_Probes.Corrosion_Probe_ID, Soil_Corrosion_Probes.Boring_Location_ID, Soil_Corrosion_Probes.Depth, " +
                                  "Soil_Corrosion_Probes.Serial_No, Soil_Corrosion_Probes.Make, Soil_Corrosion_Probes.Model, Soil_Corrosion_Probes.Installation_Method, " +
                                  "Soil_Corrosion_Probes.ER_Coupon_Material, Soil_Corrosion_Probes.dbDate_Installed, Soil_Corrosion_Probes.Divisions_Baseline, " +
                                  "Soil_Corrosion_Probes.Probe_Check_Baseline, Soil_Corrosion_Probes.Weather_Conditions " +
                                  "FROM Soil_Corrosion_Probes " +
                                  "WHERE (((Soil_Corrosion_Probes.Corrosion_Probe_ID)=\"" + _corrosion_probe_id +"\"));";

                    command = new OleDbCommand(queryString, connection);
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        // Fill in controls from the on General Page                       
                        tbCPID.Text = GetFieldDataStr("Corrosion_Probe_ID", ref reader);
                        tbBEID.Text = GetFieldDataStr("Boring_Location_ID", ref reader);
                        tbZ.Text = GetFieldDataStr("Depth", ref reader);
                        tbMake.Text = GetFieldDataStr("Make", ref reader);
                        tbCouponMatl.Text = GetFieldDataStr("ER_Coupon_Material", ref reader);
                        tbDivBaseline.Text = GetFieldDataStr("Divisions_Baseline", ref reader);
                        tbInstallationMethod.Text = GetFieldDataStr("Installation_Method", ref reader);
                        tbSN.Text = GetFieldDataStr("Serial_No", ref reader);
                        tbModel.Text = GetFieldDataStr("Model", ref reader);
                        tbDateInstalled.Text = GetFieldDataStr("dbDate_Installed", ref reader);
                        tbProbeCheckBaseline.Text = GetFieldDataStr("Probe_Check_Baseline", ref reader);
                        tbWeatherConditions.Text = GetFieldDataStr("Weather_Conditions", ref reader);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                    return false;
                }


                try
                {
                    //...............
                    // Corrosion Rate Samples Grid
                    //...............
                    queryString =
                        "SELECT [Dbf CorrosionRateSamples].Corrosion_Probe_ID, [Dbf CorrosionRateSamples].[DateStamp] FROM [Dbf CorrosionRateSamples] " +
                        "INNER JOIN Soil_Corrosion_Probes ON [Dbf CorrosionRateSamples].Corrosion_Probe_ID = Soil_Corrosion_Probes.Corrosion_Probe_ID " +
                        "WHERE (([Dbf CorrosionRateSamples].Corrosion_Probe_ID)=\"" + _corrosion_probe_id + "\");";
                    OleDbDataAdapter dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                    DataTable dTable = new DataTable();
                    //fill the DataTable
                    dAdapter.Fill(dTable);
                    DataGridView dgView = new DataGridView();
                    BindingSource bSource = new BindingSource();
                    bSource.DataSource = dTable;
                    gridCRSamples.DataSource = bSource;
                    gridCRSamples.Columns["Corrosion_Probe_ID"].Visible = false;
                }
                catch (Exception ex)
                {
                    
                }
            } // if connection!=null
            // DONE
            return true;
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
