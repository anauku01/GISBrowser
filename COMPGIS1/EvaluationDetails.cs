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
    public partial class EvaluationDetails : Form
    {
        private string _connectionstring = "";
        private int _inspectionid = 0;
        private int _height_with_additional_data = 750;
        private int _height_normal = 545;
        private int _split_height_normal = 330;
        private int _split_height_with_additional_data = 540;

        private string _method = "";

                
        //-----------------------------------------------------------------------------------
        // Constructor
        //-----------------------------------------------------------------------------------
        public EvaluationDetails(int Inspection_ID, string connectionstring)
        {
            InitializeComponent();
            _connectionstring = connectionstring;
            if (_connectionstring.Length > 0)
            {
                _inspectionid = Inspection_ID;
                LoadFormData(_inspectionid);
                this.Text = "Evaluation Details - []";
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
        public int LoadFormData(int Inspection_ID)
        {
            string reportno = "";
            string queryString = "";
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
                try
                {






                    connection.Open();
                    queryString = "SELECT [Dbf Inspections].ID, [Dbf Inspections].[Zone ID], [Dbf Inspections].[Inspection Date], [Dbf Inspections].WSLock, [Dbf Inspections].[WS Name], [Dbf Inspections].[WS Date], " +
                                  "[Dbf Inspections].[WS EndDate], [Dbf Inspections].Comments, [Dbf Inspections].[Time Last Inspection], [Dbf Inspections].[Time Next Inspection], [Dbf Inspections].[Report No], " +
                                  "[Dbf Line Section].[Line Section], [Dbf Line Section].Unit, [Dbf Line Section].[Line Section Description], [Dbf Line Section].[Line Section Ref #], [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point], " +

                                  "[Dbf Inspections].[Scheduled Work Scope], [Dbf Inspections].MaxT, [Dbf Inspections].MaxTLoc, [Dbf Inspections].MinT, [Dbf Inspections].MinTLoc, [Dbf Inspections].ServiceHours, [Dbf Inspections].Tmin, [Dbf Inspections].PreviousScope, [Dbf Inspections].CorrosionRate, " + 
                                  "[Dbf Inspections].CalculatedNSI, [Dbf Inspections].SafetyFactor, [Dbf Inspections].CPOperational, [Dbf Inspections].[Remaining Life], [Dbf Inspections].Tnom, [Dbf Inspections].[Auto Calculate], [Dbf Inspections].[Previous MinT], " + 

                                  "IIf(Not IsNull([dbf datasheet].[Report No]),IIf([Datasheet Format]=4,'RT'," +
                                  "IIf([Datasheet Format]=6,[dbf datasheet].[Method],'UT')),'VT') AS Method, " +
                                  "IIf(IsNull([Dbf Datasheet - Inspection Results].[Repair]) Or [Dbf Datasheet - Inspection Results].[Repair]=0,Null,IIf([Dbf Datasheet - Inspection Results].[Repair]=1,'No','Yes')) AS Repair, " +
                                  "[DBF WLD Component].REVIEWER1, [DBF WLD Component].DATE1, [DBF WLD Component].REVIEWER2, [DBF WLD Component].DATE2, [DBF WLD Component].REVIEW_COMPLETE, [DBF WLD Component].DATE3, " +
                                  "IIf(Not IsNull([DBF WLD Component].[Status_Inf_Req]) And [DBF WLD Component].[Status_Inf_Req]<>0,IIf([DBF WLD Component].[Status_Inf_Req]=1,'Acceptable','Rejectable'),Null) AS [Eval Disposition Component], " +
                                  "[Dbf Datasheet - Inspection Results].[Inspection Report No] AS [Fitness for Service Eval], [Dbf Datasheet - Inspection Results].REVIEWER1 AS [Eval Reviewer 1], [Dbf Datasheet - Inspection Results].DATE1 AS [Eval Date 1], " +
                                  "[Dbf Datasheet - Inspection Results].REVIEWER2 AS [Eval Reviewer 2], [Dbf Datasheet - Inspection Results].DATE2 AS [Eval Date 2], [Dbf Datasheet - Inspection Results].REVIEW_COMPLETE AS [Eval Review Complete], " +
                                  "[Dbf Datasheet - Inspection Results].DATE3 AS [Eval Date 3], IIf(Not IsNull([Dbf Datasheet - Evaluation].[STATUS]) And [Dbf Datasheet - Evaluation].[STATUS]<>0,IIf([Dbf Datasheet - Evaluation].[STATUS]=1,'Acceptable','Rejectable'),Null) AS [Eval Disposition], " +
                                  "IIf(Not IsNull([Dbf Datasheet - Evaluation].[STATUS]) And [Dbf Datasheet - Evaluation].[STATUS]=2,IIf([Dbf Datasheet - Evaluation].[COND_RPT_REQ]=True,'Yes','No'),Null) AS [INF Required] " +
                                  "FROM ((((([Dbf Line Section] RIGHT JOIN [Dbf Inspections] ON [Dbf Line Section].[Zone ID] = [Dbf Inspections].[Zone ID]) " +
                                  "LEFT JOIN [Dbf Datasheet] ON [Dbf Inspections].[Report No] = [Dbf Datasheet].[Report No]) LEFT JOIN [Dbf datasheet VT] ON [Dbf Inspections].[Report No] = [Dbf datasheet VT].SHEET_NO) " +
                                  "LEFT JOIN [Dbf Datasheet - Inspection Results] ON [Dbf Inspections].[Report No] = [Dbf Datasheet - Inspection Results].[Report No]) LEFT JOIN [DBF WLD Component] ON [Dbf Inspections].[Report No] = [DBF WLD Component].[Report No]) " +
                                  "LEFT JOIN [Dbf Datasheet - Evaluation] ON [Dbf Inspections].[Report No] = [Dbf Datasheet - Evaluation].[Report No] WHERE ((([Dbf Inspections].ID)=" + Inspection_ID + ")) ORDER BY [Dbf Inspections].[Inspection Date], " +
                                  "[Dbf Line Section].[Line Section]";
                    command = new OleDbCommand(queryString, connection);
                    try
                    {
                        reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            // Fill in header controls from the Line Section Table
                            tbLineSection.Text = GetFieldDataStr("Line Section", ref reader);
                            tbLineSectionRef.Text = GetFieldDataStr("Line Section Ref #", ref reader);
                            tbBeginPoint.Text = GetFieldDataStr("Begin Point", ref reader);
                            tbEndPoint.Text = GetFieldDataStr("End Point", ref reader);
                            tbWorkScope.Text = GetFieldDataStr("WS Name", ref reader);
                            tbUnit.Text = GetFieldDataStr("Unit", ref reader);
                            reportno = GetFieldDataStr("Report No", ref reader);
                            tbReportNo.Text = reportno;

                            // Set the form Size and visible components based upon the Method of the Evaluation
                            _method = GetFieldDataStr("Method", ref reader);
                            SetFormMode(_method);
                            // Fill out fields accoring to method
                            if (_method == "UT")
                            {
                                tbMaxT.Text = GetFieldDataStr("MaxT", ref reader);
                                tbMaxTLoc.Text = GetFieldDataStr("MaxTLoc", ref reader);
                                tbMinT.Text = GetFieldDataStr("MinT", ref reader);
                                tbMinTLoc.Text = GetFieldDataStr("MinTLoc", ref reader);
                                tbPrevMinT.Text = GetFieldDataStr("Previous MinT", ref reader);
                                tbRemLife.Text = GetFieldDataStr("Remaining Life", ref reader);
                                tbCalcNSI.Text = GetFieldDataStr("CalculatedNSI", ref reader);
                                tbNSIWS.Text = GetFieldDataStr("Scheduled Work Scope", ref reader);
                                tbInspDate.Text = GetFieldDataStr("Inspection Date", ref reader);
                                tbPrevScope.Text = GetFieldDataStr("PreviousScope", ref reader);
                                tbCorrRate.Text = GetFieldDataStr("CorrosionRate", ref reader);
                                tbSafetyFactor.Text = GetFieldDataStr("SafetyFactor", ref reader);
                                tbTmin.Text = GetFieldDataStr("Tmin", ref reader);
                                tbTnom.Text = GetFieldDataStr("Tnom", ref reader);
                                tabMain.SelectedTab = tabUTEval;
                            }
                            else
                                if (_method == "BE")
                                {
                                    tabMain.SelectedTab = tabBEEval;
                                    // Have to get the BL Datasheet with it's ID's
                                    int FieldLogID = 0;
                                    int CorrosionProbeID = 0;
                                    int BLID = 0;
                                    int GIS_Resistivity_ID = 0;
                                    int Soil_Desc_ID = 0;
                                    int DCP_Results_ID = 0;
                                    int Chem_Results_ID = 0;
                                    int Sieve_Analysis_ID = 0;
                                    queryString = "SELECT [Dbf Datasheet - BL].[Datasheet ID], [Dbf Datasheet - BL].[Report No], [Dbf Datasheet - BL].Field_Log_ID, [Dbf Datasheet - BL].Corrosion_ProbeID, "+
                                    "[Dbf Datasheet - BL].Boring_Location_ID, [Dbf Datasheet - BL].GIS_Resitivity_Values_ID, [Dbf Datasheet - BL].Soil_Description_ID, [Dbf Datasheet - BL].DCP_Results_ID, "+
                                    "[Dbf Datasheet - BL].Soil_Chemical_Results_ID, [Dbf Datasheet - BL].Soil_Sieve_Analysis_ID, [Dbf Datasheet - BL].Corrosion_Probe_ID FROM [Dbf Datasheet - BL] "+
                                    "WHERE ((([Dbf Datasheet - BL].[Report No])='" + reportno + "'))";
                                    command = new OleDbCommand(queryString, connection);
                                    try
                                    {
                                        reader = command.ExecuteReader();
                                        if (reader.Read())
                                        {
                                            FieldLogID = GetFieldDataInt("Field_Log_ID", ref reader);
                                            CorrosionProbeID = GetFieldDataInt("Corrosion_ProbeID", ref reader);
                                            BLID = GetFieldDataInt("Boring_Location_ID", ref reader);
                                            GIS_Resistivity_ID = GetFieldDataInt("GIS_Resitivity_Values_ID", ref reader);
                                            Soil_Desc_ID = GetFieldDataInt("Soil_Description_ID", ref reader);
                                            DCP_Results_ID = GetFieldDataInt("DCP_Results_ID", ref reader);
                                            Chem_Results_ID = GetFieldDataInt("Soil_Chemical_Results_ID", ref reader);
                                            Sieve_Analysis_ID = GetFieldDataInt("Soil_Sieve_Analysis_ID", ref reader);
                                        }
                                        reader.Close();

                                    }
                                    catch (Exception ex)
                                    {
                                        connection.Close();
                                        return -1;
                                    }
                                    // Fill out Results
                                    // Field Log
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
                                            tbElevation.Text = GetFieldDataStr("Elevation (ft)", ref reader);
                                            tbSampleDepthTop.Text = GetFieldDataStr("Sample_Depth_Top_FT_BGS", ref reader);
                                            tbSampleDepthBottom.Text = GetFieldDataStr("Sample_Depth_Btm_FT_BGS", ref reader);
                                            tbRecovery.Text = GetFieldDataStr("Recovery_Ft", ref reader);
                                            tbSampleDate.Text = GetFieldDataStr("SampledbDate", ref reader);
                                            tbAmbientAirTemp.Text = GetFieldDataStr("Ambient_Air_Temp", ref reader);
                                            tbPipeToSoilPotential.Text = GetFieldDataStr("Pipe_to_Soil_potential_V_SCE", ref reader);
                                        }
                                        reader.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        reader.Close();
                                    }

                                    // Soil Descriptions
                                    if (Soil_Desc_ID > 0)
                                    {
                                        queryString = "SELECT Soil_Descriptions.Soil_Description_ID, Soil_Descriptions.Field_Log_ID, Soil_Descriptions.Sample_Depth_Top_FT_BGS, " +
                                            "Soil_Descriptions.Sample_Depth_Btm_FT_BGS, Soil_Descriptions.pc_Sand, Soil_Descriptions.pc_Gravel, Soil_Descriptions.pc_Fines, " +
                                            "Soil_Descriptions.Density, Soil_Descriptions.Water_content, Soil_Descriptions.Grade, Soil_Descriptions.Color, Soil_Descriptions.Sand_Coarseness, " +
                                            "Soil_Descriptions.Sand_Shape, Soil_Descriptions.Gravel_Coarseness, Soil_Descriptions.Gravel_Shape " +
                                            "FROM Soil_Descriptions WHERE (((Soil_Descriptions.Soil_Description_ID)=" + Soil_Desc_ID + "))";
                                        command = new OleDbCommand(queryString, connection);
                                        try
                                        {
                                            reader = command.ExecuteReader();
                                            if (reader.Read())
                                            {
                                                // Fill in controls 
                                                tbpcFines.Text = GetFieldDataStr("pc_Fines", ref reader);
                                                tbpcSand.Text = GetFieldDataStr("pc_Sand", ref reader);
                                                tbpcGravel.Text = GetFieldDataStr("pc_Gravel", ref reader);
                                                tbSandCourse.Text = GetFieldDataStr("Sand_Coarseness", ref reader);
                                                tbSandShape.Text = GetFieldDataStr("Sand_Shape", ref reader);
                                                tbGravelCourse.Text = GetFieldDataStr("Gravel_Coarseness", ref reader);
                                                tbGravelShape.Text = GetFieldDataStr("Gravel_Shape", ref reader);
                                                tbSampleDensity.Text = GetFieldDataStr("Density", ref reader);
                                                tbSampleColor.Text = GetFieldDataStr("Color", ref reader);
                                                tbSampleGrade.Text = GetFieldDataStr("Grade", ref reader);                                                
                                            }
                                            reader.Close();
                                        }
                                        catch (Exception ex)
                                        {
                                            reader.Close();
                                        }
                                    }

                                    
                                    // DCP Results
                                    if (DCP_Results_ID > 0)
                                    {
                                        queryString = "SELECT Soil_DCP_Results.DCP_Results_ID, Soil_DCP_Results.Boring_Location_ID, Soil_DCP_Results.Depth, Soil_DCP_Results.Run, Soil_DCP_Results.Drive_Distance_In, "+
                                            "Soil_DCP_Results.DCP_Blows, Soil_DCP_Results.Estimated_Standard_N_Resistance_blows_ft, Soil_DCP_Results.Estimated_Cohesive_Soil_Consistency, "+
                                            "Soil_DCP_Results.Total_Drive_Distance_in, Soil_DCP_Results.Notes "+
                                            "FROM Soil_DCP_Results WHERE (((Soil_DCP_Results.DCP_Results_ID)=" + DCP_Results_ID + "))";
                                        command = new OleDbCommand(queryString, connection);
                                        try
                                        {
                                            reader = command.ExecuteReader();
                                            if (reader.Read())
                                            {
                                                // Fill in controls 
                                                tbDCPDepth.Text = GetFieldDataStr("Depth", ref reader);
                                                tbDCPRun.Text = GetFieldDataStr("Run", ref reader);
                                                tbDCPBlows.Text = GetFieldDataStr("DCP_Blows", ref reader);
                                                tbDCPDriveDistanceIn.Text = GetFieldDataStr("Drive_Distance_In", ref reader);
                                                tbDCPEstCohesive.Text = GetFieldDataStr("Estimated_Cohesive_Soil_Consistency", ref reader);
                                                tbDCPEstStd.Text = GetFieldDataStr("Estimated_Standard_N_Resistance_blows_ft", ref reader);
                                                tbDCPTotalDriveDistanceIn.Text = GetFieldDataStr("Total_Drive_Distance_in", ref reader);
                                            }
                                            reader.Close();
                                        }
                                        catch (Exception ex)
                                        {
                                            reader.Close();
                                        }
                                    }


                                    // Soil Chemical Analysis
                                    if (Chem_Results_ID > 0)
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
                                                      "WHERE (((Soil_Chemical_Results.Soil_Chemical_Results_ID)=" + Chem_Results_ID.ToString() + "))";

                                        command = new OleDbCommand(queryString, connection);
                                        reader = command.ExecuteReader();
                                        if (reader.Read())
                                        {
                                            //..........................................
                                            // Fill in grids :: Grid 1
                                            gridSoilAnalysis1.Rows.Clear();
                                            // Add Extractable Row
                                            gridSoilAnalysis1.Rows.Add(GetFieldDataStr("Extractable_Calcium_mg_kg", ref reader), 
                                                                        GetFieldDataStr("Extractable_Magnesium_mg_kg", ref reader),
                                                                        GetFieldDataStr("Extractable_Potassium_mg_kg", ref reader),
                                                                        GetFieldDataStr("Extractable_Sodium_mg_kg", ref reader),
                                                                        GetFieldDataStr("Extractable_Chloride_mg_kg", ref reader),
                                                                        GetFieldDataStr("Sulfate_Reporting_Limit", ref reader),
                                                                        GetFieldDataStr("Extractable_Sulfides_mg_kg", ref reader));
                                            // Add Result Qualifier Row
                                            gridSoilAnalysis1.Rows.Add(GetFieldDataStr("Calcium_Lab_Result_Qualifier", ref reader), 
                                                                        GetFieldDataStr("Magnesium_Lab_Result_Qualifier", ref reader),
                                                                        GetFieldDataStr("Potassium_Lab_Result_Qualifier", ref reader),
                                                                        GetFieldDataStr("Sodium_Lab_Result_Qualifier", ref reader),
                                                                        GetFieldDataStr("Chloride_Lab_Result_Qualifier", ref reader),
                                                                        GetFieldDataStr("Sulfate_Lab_Result_Qualifier", ref reader),
                                                                        GetFieldDataStr("Sulfides_Lab_Result_Qualifier", ref reader));
                                            // Reporting Limit Row
                                            gridSoilAnalysis1.Rows.Add(GetFieldDataStr("Calcium_Reporting_Limit", ref reader), 
                                                                        GetFieldDataStr("Magnesium_Reporting_Limit", ref reader),
                                                                        GetFieldDataStr("Potassium_Reporting_Limit", ref reader),
                                                                        GetFieldDataStr("Sodium_Reporting_Limit", ref reader),
                                                                        GetFieldDataStr("Chloride_Reporting_Limit", ref reader),
                                                                        GetFieldDataStr("Sulfate_Reporting_Limit", ref reader),
                                                                        GetFieldDataStr("Sulfides_Reporting_Limit", ref reader));

                                            //..........................................
                                            // Fill in grids :: gridSoilAnalysispH
                                            gridSoilAnalysispH.Rows.Clear();
                                            // Add Extractable Row
                                            gridSoilAnalysispH.Rows.Add(GetFieldDataStr("pH_SU", ref reader));
                                            // Add Results Qualifier Row
                                            gridSoilAnalysispH.Rows.Add(GetFieldDataStr("pH_SU_Lab_Result_Qualifier", ref reader));
                                            // Reporting Limit Row
                                            gridSoilAnalysispH.Rows.Add(GetFieldDataStr("pH_SU_Reporting_Limit", ref reader));
                                            //..........................................
                                            // Fill in grids :: gridSoilAnalysiORP                                            
                                            gridSoilAnalysiORP.Rows.Clear();
                                            // Add Extractable Row
                                            gridSoilAnalysiORP.Rows.Add(GetFieldDataStr("ORP_mV,", ref reader));
                                            // Add Results Qualifier Row
                                            gridSoilAnalysiORP.Rows.Add(GetFieldDataStr("ORP_mV_Lab_Result_Qualifier", ref reader));
                                            // Reporting Limit Row
                                            gridSoilAnalysiORP.Rows.Add(GetFieldDataStr("ORP_mV_Reporting_Limit", ref reader));
                                        }
                                        reader.Close();
                                    }

                                }
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        connection.Close();
                        return -1;
                    }
                    
                    // Fill Out Remaining Eval Fields from DBF WLD Component table
                    if (reportno.Length > 0)
                    {
                        queryString = "SELECT [DBF WLD Component].ID, [DBF WLD Component].[Report No], [DBF WLD Component].[Zone ID], [DBF WLD Component].COMMENTS, " +
                        "[DBF WLD Component].REVIEWER1, [DBF WLD Component].DATE1, [DBF WLD Component].REVIEWER2, [DBF WLD Component].DATE2, [DBF WLD Component].REVIEW_COMPLETE, " +
                        "[DBF WLD Component].DATE3, [DBF WLD Component].STATUS, [DBF WLD Component].Title1, [DBF WLD Component].Title2, [DBF WLD Component].Title3, [DBF WLD Component].Pre_Exam, " +
                        "[DBF WLD Component].Numbers, [DBF WLD Component].Mec_Analysis, [DBF WLD Component].Final_Disposition, [DBF WLD Component].Status_Inf_Req, [DBF WLD Component].PC_No, " +
                        "[DBF WLD Component].Cond_Rpt_Req, [DBF WLD Component].Title, [DBF WLD Component].Soil_Results_Comments, [DBF WLD Component].Sample_Comments, [DBF WLD Component].Evaluation_Comments, " +
                        "[DBF WLD Component].Apply_Soil_Results FROM [DBF WLD Component] " +
                        "WHERE ((([DBF WLD Component].[Report No])='" + reportno + "'))";
                        command = new OleDbCommand(queryString, connection);
                        try
                        {
                            reader = command.ExecuteReader();
                            if (reader.Read())
                            {
                                // Fill in header controls from the Line Section Table
                                tbReviewer1.Text = GetFieldDataStr("REVIEWER1", ref reader);
                                tbReviewer2.Text = GetFieldDataStr("REVIEWER2", ref reader);
                                tbReviewComplete.Text = GetFieldDataStr("REVIEW_COMPLETE", ref reader);
                                tbReview1Title.Text = GetFieldDataStr("Title1", ref reader);
                                tbReview2Title.Text = GetFieldDataStr("Title2", ref reader);
                                tbReviewCompleteTitle.Text = GetFieldDataStr("Title3", ref reader);
                                tbReview1Date.Text = GetFieldDataStr("DATE1", ref reader);
                                tbReview2Date.Text = GetFieldDataStr("DATE1", ref reader);
                                tbReviewCompleteDate.Text = GetFieldDataStr("DATE3", ref reader);
                                // Disposition
                                int disp = Convert.ToInt32(GetFieldDataStr("Status_Inf_Req", ref reader));
                                cbEvalDispAccept.Checked = (disp == 1);
                                cbEvalDispReject.Checked = (disp == 0);
                                // Compared with Previous
                                int pre_exam = Convert.ToInt32(GetFieldDataStr("Pre_Exam", ref reader));
                                cbComparedToPrevNA.Checked = (pre_exam == 1);
                                cbComparedToPrevNo.Checked = (pre_exam == 2);
                                cbComparedToPrevYes.Checked = (pre_exam == 3);
                                // Fracture Mechanics
                                cbFractureYes.Checked = (Convert.ToInt32(GetFieldDataStr("Mec_Analysis", ref reader)) != 0);
                                cbFractureNo.Checked = !cbFractureYes.Checked;
                                tbComments.Text = GetFieldDataStr("Final_Disposition", ref reader);
                                tbMainComments.Text = GetFieldDataStr("Comments", ref reader);
                            }
                            reader.Close();

                        }
                        catch (Exception ex)
                        {
                            connection.Close();
                            return -1;
                        }
                    }

                } // try...
                finally
                {
                    connection.Close();
                }
            return 0;
        } // LoadFormData


        //-----------------------------------------------------------------------------------
        // Sets the Form Usage Mode depending on the Method of the Evaluation
        //-----------------------------------------------------------------------------------
        private void SetFormMode(string Method)
        {
            if ((Method == "VT")||(Method == "RT")) // VT and RT have no additional Evaluation data unlike UT, BE
            {
                scMain.Panel1Collapsed = true;
                scMain.Height = _split_height_normal;
                Height = _height_normal;
            }
            else
            {
                scMain.Panel1Collapsed = false;
                scMain.Height = _split_height_with_additional_data;
                Height = _height_with_additional_data;
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
