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
using System.Drawing.Imaging;

namespace COMPGIS1
{
    public partial class BPChemicalAnalysisDetails : Form
    {

        private string _connectionstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Iddeal_Server_Ver8\\Iddeal_DE1\\Data\\Comp_BP_Data_Lasalle.mdb;Persist Security Info=True;Password=waiheke;User ID=developer;Jet OLEDB:System database=C:\\Iddeal_Server_Ver8\\V08.mdw;Jet OLEDB:Database Password=waiheke";
        private int _soil_chemical_results_id;
        private bool _showradarplot = false;

        public BPChemicalAnalysisDetails(int Soil_Chemical_Results_ID, string connectionstring)
        {
            InitializeComponent();
            if (Soil_Chemical_Results_ID > 0)
            {
                _soil_chemical_results_id = Soil_Chemical_Results_ID;
                _connectionstring = connectionstring;
                LoadFormData(_soil_chemical_results_id);
                RadarPlotGenerator rpg = new RadarPlotGenerator(chart1, "", "", "",Soil_Chemical_Results_ID,_connectionstring);
                _showradarplot = rpg.plotcomplete;
                chart1.Visible = _showradarplot;
                btnSaveImg.Visible = _showradarplot;
                this.Text = "Chemical Analysis Details";
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
        public int LoadFormData(int Soil_Chemical_Results_ID)
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
                    queryString = "SELECT Soil_Chemical_Results.Soil_Chemical_Results_ID, Soil_Chemical_Results.Boring_Location_ID, Soil_Chemical_Results.Z, " +
                                  "Soil_Chemical_Results.Sample_Depth_Top_FT_BGS, Soil_Chemical_Results.Sample_Depth_Btm_FT_BGS, Soil_Chemical_Results.CollectionDate, " +
                                  "Soil_Chemical_Results.AnalysisDate, Soil_Chemical_Results.LabName, Soil_Chemical_Results.AnalysisMethod, Soil_Chemical_Results.Extractable_Calcium_mg_kg, " +
                                  "Soil_Chemical_Results.Calcium_Lab_Result_Qualifier, Soil_Chemical_Results.Calcium_Reporting_Limit, Soil_Chemical_Results.Extractable_Magnesium_mg_kg, " +
                                  "Soil_Chemical_Results.Magnesium_Lab_Result_Qualifier, Soil_Chemical_Results.Magnesium_Reporting_Limit, Soil_Chemical_Results.Extractable_Potassium_mg_kg, " +
                                  "Soil_Chemical_Results.Potassium_Lab_Result_Qualifier, Soil_Chemical_Results.Potassium_Reporting_Limit, Soil_Chemical_Results.Extractable_Sodium_mg_kg, " +
                                  "Soil_Chemical_Results.Sodium_Lab_Result_Qualifier, Soil_Chemical_Results.Sodium_Reporting_Limit, Soil_Chemical_Results.Extractable_Chloride_mg_kg, " +
                                  "Soil_Chemical_Results.Chloride_Lab_Result_Qualifier, Soil_Chemical_Results.Chloride_Reporting_Limit, Soil_Chemical_Results.Extractable_Sulfate_mg_kg, " +
                                  "Soil_Chemical_Results.Sulfate_Lab_Result_Qualifier, Soil_Chemical_Results.Sulfate_Reporting_Limit, Soil_Chemical_Results.Extractable_Sulfides_mg_kg, " +
                                  "Soil_Chemical_Results.Sulfides_Lab_Result_Qualifier, Soil_Chemical_Results.Sulfides_Reporting_Limit, Soil_Chemical_Results.Sulfides_AWWA_Rating, " +
                                  "Soil_Chemical_Results.pH_SU, Soil_Chemical_Results.pH_SU_Lab_Result_Qualifier, Soil_Chemical_Results.pH_SU_Reporting_Limit, " +
                                  "Soil_Chemical_Results.pH_SU_AWWA_Rating, Soil_Chemical_Results.ORP_mV, Soil_Chemical_Results.ORP_mV_Lab_Result_Qualifier, Soil_Chemical_Results.ORP_mV_Reporting_Limit, " +
                                  "Soil_Chemical_Results.ORP_mV_AWWA_Rating, Soil_Chemical_Results.Resistivity_Ohm_cm, Soil_Chemical_Results.Resistivity_Ohm_cm_Lab_Result_Qualifier, " +
                                  "Soil_Chemical_Results.Resistivity_Ohm_cm_Reporting_Limit, Soil_Chemical_Results.Resistivity_Ohm_cm_AWWA_Rating, Soil_Chemical_Results.Resistivity_Ohm_cm_Classification, " +
                                  "Soil_Chemical_Results.pcMoisture_wet_wt, Soil_Chemical_Results.pcMoisture_wet_wt_Lab_Result_Qualifier, Soil_Chemical_Results.pcMoisture_wet_wt_Reporting_Limit, " +
                                  "Soil_Chemical_Results.pcMoisture_wet_wt_AWWA_Rating, Soil_Chemical_Results.pcSolids_wet_wt, Soil_Chemical_Results.pcSolids_wet_wt_Lab_Result_Qualifier, " +
                                  "Soil_Chemical_Results.pcSolids_wet_wt_Reporting_Limit, Soil_Chemical_Results.pcSolids_wet_wt_AWWA_Rating, Soil_Chemical_Results.[Pipe-to-Soil_Potential_mV], " +
                                  "Soil_Chemical_Results.PtS_Potential_mV_Lab_Result_Qualifier, Soil_Chemical_Results.PtS_Potential_mV_Reporting_Limit, Soil_Chemical_Results.PtS_Potential_mV_AWWA_Rating, " +
                                  "Soil_Chemical_Results.AWWA_Point_Total, Soil_Chemical_Results.SRB_Found, Soil_Chemical_Results.SRB_Amount, Soil_Chemical_Results.PolarPlotImage FROM Soil_Chemical_Results " +
                                  "WHERE (((Soil_Chemical_Results.Soil_Chemical_Results_ID)="+ Soil_Chemical_Results_ID.ToString() +"));";

                    command = new OleDbCommand(queryString, connection);
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        // Fill in controls from the on General Page
                        tbBEID.Text = GetFieldDataStr("Boring_Location_ID", ref reader);
                        tbZ.Text = GetFieldDataStr("Z", ref reader);
                        tbSampleDepthTop.Text = GetFieldDataStr("Sample_Depth_Top_FT_BGS", ref reader);
                        tbSampleDepthBottom.Text = GetFieldDataStr("Sample_Depth_Btm_FT_BGS", ref reader);
                        tbAnalysisDate.Text = GetFieldDataStr("AnalysisDate", ref reader);
                        tbCollectionDate.Text = GetFieldDataStr("CollectionDate", ref reader);
                        tbLabName.Text = GetFieldDataStr("LabName", ref reader);
                        tbAnalysisMethod.Text = GetFieldDataStr("AnalysisMethod", ref reader);
                        cbSRBFound.Checked =(Convert.ToBoolean(GetFieldDataStr("SRB_Found", ref reader)));
                        tbSRBAmount.Text = GetFieldDataStr("SRB_Amount", ref reader);
                        tbAWWAPointTotal.Text = GetFieldDataStr("AWWA_Point_Total", ref reader);

                        // Results Page
                        // Ca
                        tbExtCa.Text = GetFieldDataStr("Extractable_Calcium_mg_kg", ref reader);
                        tbLRQCa.Text = GetFieldDataStr("Calcium_Lab_Result_Qualifier", ref reader);
                        tbRLCa.Text = GetFieldDataStr("Calcium_Reporting_Limit", ref reader);
                        // Mg
                        tbExtMg.Text = GetFieldDataStr("Extractable_Magnesium_mg_kg", ref reader);
                        tbLRQMg.Text = GetFieldDataStr("Magnesium_Lab_Result_Qualifier", ref reader);
                        tbRLMg.Text = GetFieldDataStr("Magnesium_Reporting_Limit", ref reader);
                        // Potassium
                        tbExtK.Text = GetFieldDataStr("Extractable_Potassium_mg_kg", ref reader);
                        tbLRQK.Text = GetFieldDataStr("Potassium_Lab_Result_Qualifier", ref reader);
                        tbRLK.Text = GetFieldDataStr("Potassium_Reporting_Limit", ref reader);
                        // Sodium
                        tbExtNa.Text = GetFieldDataStr("Extractable_Sodium_mg_kg", ref reader);
                        tbLRQNa.Text = GetFieldDataStr("Sodium_Lab_Result_Qualifier", ref reader);
                        tbRLNa.Text = GetFieldDataStr("Sodium_Reporting_Limit", ref reader);
                        // Chloride
                        tbExtCl.Text = GetFieldDataStr("Extractable_Chloride_mg_kg", ref reader);
                        tbLRQCl.Text = GetFieldDataStr("Chloride_Lab_Result_Qualifier", ref reader);
                        tbRLCl.Text = GetFieldDataStr("Chloride_Reporting_Limit", ref reader);
                        // Sulphate
                        tbExtS042.Text = GetFieldDataStr("Extractable_Sulfate_mg_kg", ref reader);
                        tbLRQS042.Text = GetFieldDataStr("Sulfate_Lab_Result_Qualifier", ref reader);
                        tbRLS042.Text = GetFieldDataStr("Sulfate_Reporting_Limit", ref reader);
                        // Sulphides
                        tbExtS2.Text = GetFieldDataStr("Extractable_Sulfides_mg_kg", ref reader);
                        tbLRQS2.Text = GetFieldDataStr("Sulfides_Lab_Result_Qualifier", ref reader);
                        tbRLS2.Text = GetFieldDataStr("Sulfides_Reporting_Limit", ref reader);
                        // pH SU
                        tbpHSU.Text = GetFieldDataStr("pH_SU", ref reader);
                        tbRLpHSU.Text = GetFieldDataStr("pH_SU_Reporting_Limit", ref reader);
                        tbAWWApHSU.Text = GetFieldDataStr("pH_SU_AWWA_Rating", ref reader);
                        tbLRQpHSU.Text = GetFieldDataStr("pH_SU_Lab_Result_Qualifier", ref reader);
                        // ORP
                        tbORPmv.Text = GetFieldDataStr("ORP_mV", ref reader);
                        tbRLORPmv.Text = GetFieldDataStr("ORP_mV_Reporting_Limit", ref reader);
                        tbAWWAORPmv.Text = GetFieldDataStr("ORP_mV_AWWA_Rating", ref reader);
                        tbLRQORPmv.Text = GetFieldDataStr("ORP_mV_Lab_Result_Qualifier", ref reader);
                        // Resistivity
                        tbResistivity.Text = GetFieldDataStr("Resistivity_Ohm_cm", ref reader);
                        tbRLResistivity.Text = GetFieldDataStr("Resistivity_Ohm_cm_Lab_Result_Qualifier", ref reader);
                        tbAWWAResistivity.Text = GetFieldDataStr("Resistivity_Ohm_cm_AWWA_Rating", ref reader);
                        tbRLResistivity.Text = GetFieldDataStr("Resistivity_Ohm_cm_Reporting_Limit", ref reader);
                        // Moisture
                        tbMoistureWet.Text = GetFieldDataStr("pcMoisture_wet_wt", ref reader);
                        tbLRQMoisture.Text = GetFieldDataStr("pcMoisture_wet_wt_Lab_Result_Qualifier", ref reader);
                        tbRLMoisture.Text = GetFieldDataStr("pcMoisture_wet_wt_Reporting_Limit", ref reader);
                        tbAWWAMoisture.Text = GetFieldDataStr("pcMoisture_wet_wt_AWWA_Rating", ref reader);
                        // Solids
                        tbSolidsWetwt.Text = GetFieldDataStr("pcSolids_wet_wt", ref reader);
                        tbLRQSolids.Text = GetFieldDataStr("pcSolids_wet_wt_Lab_Result_Qualifier", ref reader);
                        tbRLSolids.Text = GetFieldDataStr("pcSolids_wet_wt_Reporting_Limit", ref reader);
                        tbAWWASolids.Text = GetFieldDataStr("pcSolids_wet_wt_AWWA_Rating", ref reader);
                        // Pipe To Soil Potential
                        tbPTSPotential.Text = GetFieldDataStr("Pipe-to-Soil_Potential_mV", ref reader);
                        tbLRQPtSPotential.Text = GetFieldDataStr("PtS_Potential_mV_Lab_Result_Qualifier", ref reader);
                        tbRLPtSPotential.Text = GetFieldDataStr("PtS_Potential_mV_Reporting_Limit", ref reader);
                        tbAWWAPtSPotential.Text = GetFieldDataStr("PtS_Potential_mV_AWWA_Rating", ref reader);

                    }
                    reader.Close();

                }
                catch (Exception ex)
                {
                    connection.Close();
                    return -1;
                }
                //connection.Close();
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

        //-----------------------------------------------------------------------------------
        // User Clicked Save Image
        //-----------------------------------------------------------------------------------
        private void btnSaveImg_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "jpg image (*.jpg)|.jpg";
            saveFileDialog1.FilterIndex = 0;
            saveFileDialog1.RestoreDirectory = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
//                chart1.SaveImage(saveFileDialog1.OpenFile(),System.Drawing.Image);
            }
        } 



    }
}