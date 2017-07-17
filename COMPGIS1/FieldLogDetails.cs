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
    public partial class FieldLogDetails : Form
    {

        private string _connectionstring = "";
        private int _fieldlogid = 0;


        //-----------------------------------------------------------------------------------
        // Contructor
        //-----------------------------------------------------------------------------------
        public FieldLogDetails(int fieldlogid, string connectionstring)
        {
            InitializeComponent();
            _connectionstring = connectionstring;
            if (_connectionstring.Length > 0)
            {
                _fieldlogid = fieldlogid;
                if (LoadFormData(_fieldlogid) == 0)
                {
                    this.Text = string.Format("Field Log Details - [{0}] top: [{1}] ", tbBEID.Text, tbSampleDepthTop.Text);
                }
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
        public int LoadFormData(int FieldLogID)
        {
            string queryString = "";
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);

            if (connection != null)
            {
                connection.Open();
                queryString = "SELECT Soil_Field_Log.Field_Log_ID, Soil_Field_Log.Boring_Location_ID, Soil_Field_Log.SampledbDate, Soil_Field_Log.[Elevation (ft)], " +
                "Soil_Field_Log.Sample_Depth_Top_FT_BGS, Soil_Field_Log.Sample_Depth_Btm_FT_BGS, Soil_Field_Log.Recovery_Ft, Soil_Field_Log.Boring_Description, " + 
                "Soil_Field_Log.Pipe_to_Soil_potential_V_SCE, Soil_Field_Log.Groundwater_Level, Soil_Field_Log.Ambient_Air_Temp " + 
                "FROM Soil_Field_Log " + 
                "WHERE (((Soil_Field_Log.Field_Log_ID)=" + FieldLogID.ToString() + "))";
                command = new OleDbCommand(queryString, connection);
                try
                {
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        // Fill in controls 
                        tbBEID.Text = GetFieldDataStr("Boring_Location_ID", ref reader);
                        tbSampleDate.Text = GetFieldDataStr("SampledbDate", ref reader);
                        tbElevation.Text = GetFieldDataStr("Elevation (ft)", ref reader);
                        tbSampleDepthTop.Text = GetFieldDataStr("Sample_Depth_Top_FT_BGS", ref reader);
                        tbSampleDepthBottom.Text = GetFieldDataStr("Sample_Depth_Btm_FT_BGS", ref reader);
                        tbRecovery.Text = GetFieldDataStr("Recovery_Ft", ref reader);
                        tbGroundwaterLevel.Text = GetFieldDataStr("Groundwater_Level", ref reader);
                        tbPipeToSoilPotential.Text = GetFieldDataStr("Pipe_to_Soil_potential_V_SCE", ref reader);
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
            // DONE
            return 0;
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
