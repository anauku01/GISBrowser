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
    public partial class BPSieveAnalysisDetails : Form
    {


        private string _connectionstring = "";
        private int _sieve_analysis_id = 0;

        //-----------------------------------------------------------------------------------
        // Contructor
        //-----------------------------------------------------------------------------------
        public BPSieveAnalysisDetails(int sieve_analysis_id, string connectionstring)
        {
            InitializeComponent();
            _connectionstring = connectionstring;
            if (_connectionstring.Length > 0)
            {
                _sieve_analysis_id = sieve_analysis_id;
                if (LoadFormData(_sieve_analysis_id) == 0)
                {
                    this.Text = string.Format("Sieve Analysis Details - [{0}] top: [{1}] ", tbBEID.Text, tbSampleDepthTop.Text);
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
        public int LoadFormData(int ID)
        {
            string queryString = "";
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);

            if (connection != null)
            {
                connection.Open();
                queryString = "SELECT Soil_Sieve_Analysis.Soil_Sieve_Analysis_ID, Soil_Sieve_Analysis.Boring_Location_ID, Soil_Sieve_Analysis.Z, " +
                                "Soil_Sieve_Analysis.Sample_Depth_Top_FT_BGS, Soil_Sieve_Analysis.Sample_Depth_Btm_FT_BGS, Soil_Sieve_Analysis.Sieve_3_inch, " +
                                "Soil_Sieve_Analysis.Sieve_2_inch, Soil_Sieve_Analysis.Sieve_1_1_2_inch, Soil_Sieve_Analysis.Sieve_3_4_inch, " +
                                "Soil_Sieve_Analysis.Sieve_3_8_inch, Soil_Sieve_Analysis.Sieve_Num_4, Soil_Sieve_Analysis.Sieve_Num_10, " +
                                "Soil_Sieve_Analysis.Sieve_Num_20, Soil_Sieve_Analysis.Sieve_Num_40, Soil_Sieve_Analysis.Sieve_Num_60, " +
                                "Soil_Sieve_Analysis.Sieve_Num_100, Soil_Sieve_Analysis.Sieve_075, Soil_Sieve_Analysis.Sieve_023, Soil_Sieve_Analysis.Sieve_0136, " +
                                "Soil_Sieve_Analysis.Sieve_0098, Soil_Sieve_Analysis.Sieve_007, Soil_Sieve_Analysis.Sieve_0035, Soil_Sieve_Analysis.Sieve_0015, " +
                                "Soil_Sieve_Analysis.Gravel_pc, Soil_Sieve_Analysis.Sand_Course_pc, Soil_Sieve_Analysis.SandMedium_pc, Soil_Sieve_Analysis.SandFine_pc, " +
                                "Soil_Sieve_Analysis.Silt_pc, Soil_Sieve_Analysis.Clay_colloids_pc, Soil_Sieve_Analysis.Sieve_1_inch, Soil_Sieve_Analysis.Sieve_036 " +
                                "FROM Soil_Sieve_Analysis WHERE (((Soil_Sieve_Analysis.Soil_Sieve_Analysis_ID)=" + ID.ToString() + "))";

                command = new OleDbCommand(queryString, connection);
                try
                {
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        // Fill in controls 
                        tbBEID.Text = GetFieldDataStr("Boring_Location_ID", ref reader);
                        tbSampleDepthTop.Text = GetFieldDataStr("Sample_Depth_Top_FT_BGS", ref reader);
                        tbSampleDepthBottom.Text = GetFieldDataStr("Sample_Depth_Btm_FT_BGS", ref reader);
                        tb3inch.Text = GetFieldDataStr("Sieve_3_inch", ref reader);
                        tb2inch.Text = GetFieldDataStr("Sieve_2_inch", ref reader);
                        tb1_12inch.Text = GetFieldDataStr("Sieve_1_1_2_inch", ref reader);
                        tb34inch.Text = GetFieldDataStr("Sieve_3_4_inch", ref reader);
                        tb38inch.Text = GetFieldDataStr("Sieve_3_8_inch", ref reader);
                        tbNum4.Text = GetFieldDataStr("Sieve_Num_4", ref reader);
                        tbNum10.Text = GetFieldDataStr("Sieve_Num_10", ref reader);
                        tbNum20.Text = GetFieldDataStr("Sieve_Num_20", ref reader);
                        tbNum40.Text = GetFieldDataStr("Sieve_Num_40", ref reader);
                        tbNum60.Text = GetFieldDataStr("Sieve_Num_60", ref reader);
                        tbNum100.Text = GetFieldDataStr("Sieve_Num_100", ref reader);
                        tb075.Text = GetFieldDataStr("Sieve_075", ref reader);
                        tb023.Text = GetFieldDataStr("Sieve_023", ref reader);
                        tb0136.Text = GetFieldDataStr("Sieve_0136", ref reader);
                        tb0098.Text = GetFieldDataStr("Sieve_0098", ref reader);
                        tb007.Text = GetFieldDataStr("Sieve_007", ref reader);
                        tb0035.Text = GetFieldDataStr("Sieve_0035", ref reader);
                        tb0015.Text = GetFieldDataStr("Sieve_0015", ref reader);
                        tbGravelpc.Text = GetFieldDataStr("Gravel_pc", ref reader);
                        tbSandCoursepc.Text = GetFieldDataStr("Sand_Course_pc", ref reader);
                        tbSandMediumpc.Text = GetFieldDataStr("SandMedium_pc", ref reader);
                        tbSandFinepc.Text = GetFieldDataStr("SandFine_pc", ref reader);
                        tbClaypc.Text = GetFieldDataStr("Clay_colloids_pc", ref reader);
                        tbSiltpc.Text = GetFieldDataStr("Silt_pc", ref reader);
                        tb1inch.Text = GetFieldDataStr("Sieve_1_inch", ref reader);
                        tb036.Text = GetFieldDataStr("Sieve_036", ref reader);
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








