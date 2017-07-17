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
    public partial class SWFlowSegmentDetails : Form
    {

        private string _connectionstring = "";
        private int _zoneid = 0;
        private int _lineid;
        private int _beginpoint = -1;
        private int _endpoint = -1;
        private int _suscrisktotal = 0;
        private int _consrisktotal = 0;
        private int _soilrisktotal = 0;
        private int _suscsoilrisktotal = 0;
        private bool _allow_open_lines = false;
        private Color _rr_low_color = Color.LimeGreen;
        private Color _rr_low_medium_color = Color.Yellow;
        private Color _rr_medium_color = Color.Coral;
        private Color _rr_medium_high_color = Color.Fuchsia;
        private Color _rr_high_color = Color.Red;

        //-----------------------------------------------------------------------------------
        // Contructor
        //-----------------------------------------------------------------------------------
        public SWFlowSegmentDetails(int Zone_ID, string connectionstring, bool allow_open_lines)
        {
            InitializeComponent();
            _connectionstring = connectionstring;
            _allow_open_lines = allow_open_lines;
            if (_connectionstring.Length > 0)
            {
                _zoneid = Zone_ID;
                LoadFormData(_zoneid);
                this.Text = "Flow Segment Details - [" + tbFlowSegment.Text + "]";
            }
        }


        //-----------------------------------------------------------------------------------
        // Get Double Field Data from the Reader
        //-----------------------------------------------------------------------------------
        private bool GetFieldDataDouble(string DataFieldName, ref OleDbDataReader rdr, ref double dblvalue)
        {
            dblvalue = 0;
            try
            {
                int idx = rdr.GetOrdinal(DataFieldName);
                if (idx >= 0)
                {
                    string str = rdr[idx].ToString();
                    if (str.Length > 0)
                    {
                        dblvalue = Convert.ToDouble(str);
                        return true;
                    }                    
                }                    
            }
            catch (Exception ex)
            {
                return false;
            }
            return false;
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
            int contentlookupid = -1;
            int lineclasslookupid = -1;
            string queryString = "";
            string queryStringLine = "";
            btnShowLine.Visible = _allow_open_lines;
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
                try
                {
                    connection.Open();
                    string ZoneStr = Zone_ID.ToString();
                    queryString = "SELECT [Dbf Zone].[Zone ID], [Dbf Zone].[Flow Segment], [Dbf Zone].[Flow Segment Ref #], [Dbf Zone].[Flow Segment Type], [Dbf Zone].[Flow Segment Description], [Dbf Zone].Node, [Dbf Zone].[Flow (gpm)], [Dbf Zone].[Flow Density], [Dbf Zone].[Nominal Pipe Size], [Dbf Zone].[Pipe OD], [Dbf Zone].[Pipe ID], [Dbf Zone].Area, [Dbf Zone].Velocity, [Dbf Zone].Comments, [Dbf Zone].Scheduled, [Dbf Zone].PID, [Dbf Zone].Iso, [Dbf Zone].PRA, [Dbf Zone].LCO, [Dbf Zone].Accessablility, [Dbf Zone].LineSize, [Dbf Zone].Building, [Dbf Zone].Elevation, [Dbf Zone].Row, [Dbf Zone].Col, [Dbf Zone].SR, [Dbf Zone].Schedule, [Dbf Zone].Category, [Dbf Zone].ChemTF, [Dbf Zone].Online, [Dbf Zone].SRB, [Dbf Zone].TAB, [Dbf Zone].APB, [Dbf Zone].IRB, [Dbf Zone].Cl, [Dbf Zone].SO3, [Dbf Zone].HCO2, [Dbf Zone].CO2, [Dbf Zone].Comments2, [Dbf Zone].InspLocation, [Dbf Zone].Drawing1, [Dbf Zone].Drawing2, [Dbf Zone].Unit, [Dbf Zone].SubSystemID, [Dbf Zone].LineID, [Dbf Zone].SystemID, [Dbf Zone].[Tnom Main], [Dbf Zone].BeginPoint, [Dbf Zone].EndPoint, [Dbf Zone].MapFeatureID FROM [Dbf Zone] WHERE ((([Dbf Zone].[Zone ID])=" + ZoneStr + "))"
                    command = new OleDbCommand(queryString, connection);
                    try
                    {
                        reader = command.ExecuteReader();
                        if (reader.Read())
                        {
                            // Fill in controls from the Line Section Table
                            tbFlowSegment.Text = GetFieldDataStr("Flow Segment", ref reader);
                            tbBeginPoint.Text = GetFieldDataStr("BeginPoint", ref reader);
                            tbEndPoint.Text = GetFieldDataStr("EndPoint", ref reader);
                            _beginpoint = Convert.ToInt32(tbBeginPoint.Text);
                            _endpoint = Convert.ToInt32(tbEndPoint.Text);
                            tbUnit.Text = GetFieldDataStr("Unit", ref reader);
                            tbFlowSegmentType.Text = GetFieldDataStr("Flow Segment Type", ref reader);
                            tbFlowSegmentRef.Text = GetFieldDataStr("Flow Segment Ref #", ref reader);
                            tbPID.Text = GetFieldDataStr("PID", ref reader);
                            tbISO.Text = GetFieldDataStr("Iso", ref reader);
                            tbAccessibility.Text = GetFieldDataStr("Accessablility", ref reader);
                            tbFlowSegmentDesc.Text = GetFieldDataStr("Flow Segment Description", ref reader);
                            tbRow.Text = GetFieldDataStr("Row", ref reader);
                            tbCol.Text = GetFieldDataStr("Col", ref reader);
                            tbCategory.Text = GetFieldDataStr("Category", ref reader);
                            tbOnline.Text = GetFieldDataStr("Online", ref reader);
                            tbInspectionLocation.Text = GetFieldDataStr("InspLocation", ref reader);


                            tbLocation.Text = GetFieldDataStr("Location", ref reader);
                            tbLocationComments.Text = GetFieldDataStr("Location Comments", ref reader);
                            tbNode.Text = GetFieldDataStr("Node", ref reader);
                            tbLoop.Text = GetFieldDataStr("Loop", ref reader);
                            tbBuilding.Text = GetFieldDataStr("Building", ref reader);
                            tbBuildingZone.Text = GetFieldDataStr("Building Zone", ref reader);
                            tbElevation.Text = GetFieldDataStr("Elevation", ref reader);
                            tbRoom.Text = GetFieldDataStr("Room", ref reader);
                            tbArea.Text = GetFieldDataStr("Area", ref reader);
                            tbAreaDescription.Text = GetFieldDataStr("Area Description", ref reader);
                            tbCoordinates.Text = GetFieldDataStr("Coordinates", ref reader);

                            // Boolean Fields
                            cbSafetyRelated.Checked = (Convert.ToBoolean(GetFieldDataStr("Safety Related", ref reader)));

                            // Lookup Fields
                            _lineid = Convert.ToInt32(GetFieldDataStr("LineID", ref reader));
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        connection.Close();
                        return -1;
                    }


                    //...............
                    // Get the Parent Line information
                    //...............
                    if (_lineid > 0)
                    {
                        queryString =
                            "SELECT [Dbf Lines].LineID, [Dbf Lines].Line, [Dbf Lines].Unit, [Dbf Lines].LineNumber, [Dbf Lines].LineDescription, [Dbf Lines].LineClass, [Dbf Lines].Line, [Dbf Lines].LinePhase, [Dbf Lines].LineSafetyGrade, [Dbf Lines].LineSize, [Dbf Lines].BoreClass, [Dbf Lines].Category, [Dbf Lines].Criteria, [Dbf Lines].BasisComments, [Dbf Lines].FailureConsequence, [Dbf Lines].[Level of Susceptability], [Dbf Lines].Temperature, [Dbf Lines].[Line Content], [Dbf Lines].[Operating Time], [Dbf Lines].EvalComments, [Dbf Lines].SNMComments, [Dbf Lines].References, [Dbf Lines].PlantExperience, [Dbf Lines].PlantExperienceDescription, [Dbf Lines].IndustryExperience, [Dbf Lines].IndustryExperienceDescription, [Dbf Lines].CheckWorksLineName, [Dbf Lines].CheckWorksLineNumber, [Dbf Lines].[Failure Affects Safe Shutdown or CDF], [Dbf Lines].[Radiological Content], [Dbf Lines].Inservice, [Dbf Lines].[Tank Type], [Dbf Lines].[Begin Point], [Dbf Lines].[End Point], [Dbf Lines].CompanyID, [Dbf Lines].GUID, [Dbf Lines].OtherSpec, [Dbf Lines].Operation, [Dbf Lines].Velocity, [Dbf Lines].NSIAC, [Dbf Lines].LicenseRenewalCommitment, [Dbf Lines].CommitmentNumber, [Dbf Lines].LRCRemarks, [Dbf Lines].IsGWPILine, [Dbf Lines].[Total Length] " +
                            "FROM [Dbf Lines] WHERE ([Dbf Lines].LineID=" + _lineid + ")";
                        command = new OleDbCommand(queryString, connection);
                        try
                        {
                            reader = command.ExecuteReader();
                            if (reader.Read())
                            {
                                tbLine.Text = GetFieldDataStr("Line", ref reader);
                                // Get Lookup IDs
                                lineclasslookupid = Convert.ToInt32(GetFieldDataStr("LineClass", ref reader));
                                contentlookupid = Convert.ToInt32(GetFieldDataStr("Line Content", ref reader));
                            }
                            reader.Close();
                        }
                        catch (Exception ex)
                        {
                            connection.Close();
                            return -1;
                        }

                        // Add checked values to the data grid
                        // Clear the grids
                        lvCons.Items.Clear();
                        lvSoil.Items.Clear();
                        // Boolean Fields
                        // Add the Risk Ranking
                        if ((_beginpoint >= 0) && (_endpoint > 0))
                        {
                            queryStringLine =
                                "SELECT [Dbf Susceptibility LS].LineID, [Dbf Susceptibility LS].[Begin Point], [Dbf Susceptibility LS].[End Point], [Dbf Susceptibility LS].Under_Road, [Dbf Susceptibility LS].Under_Bldg, [Dbf Susceptibility LS].Under_RR, [Dbf Susceptibility LS].Under_Tower_Footer, [Dbf Susceptibility LS].Under_Trans_Line, [Dbf Susceptibility LS].Over_Under_River, [Dbf Susceptibility LS].In_To_Out_Of_Bldg, [Dbf Susceptibility LS].In_Wall, [Dbf Susceptibility LS].Underground_Tee, [Dbf Susceptibility LS].Replaced, [Dbf Susceptibility LS].Mtl_Chg, [Dbf Susceptibility LS].Inspected, [Dbf Susceptibility LS].Above_Groundwater_Level, [Dbf Susceptibility LS].Earthen_Fill_Material_Change, [Dbf Susceptibility LS].Pipe, [Dbf Susceptibility LS].PipeAge, [Dbf Susceptibility LS].Leak_History, [Dbf Susceptibility LS].NotCoated, [Dbf Susceptibility LS].SoilOutOfSpec, [Dbf Susceptibility LS].GroundSettlement, [Dbf Susceptibility LS].PipesInArea, [Dbf Susceptibility LS].SteamLine, [Dbf Susceptibility LS].WallThinner, [Dbf Susceptibility LS].Potential, [Dbf Susceptibility LS].Uncorrected, [Dbf Susceptibility LS].BackfillUnacceptable, [Dbf Susceptibility LS].RectifierOperational, [Dbf Susceptibility LS].[Internal Erosion Corrosion], [Dbf Susceptibility LS].[Soil Characteristics Unknown], [Dbf Susceptibility LS].[Recorded Transient Not Corrected], [Dbf Susceptibility LS].[Within 10 ft Transmission Line Footer], [Dbf Susceptibility LS].[No Coating Inspection Performed], [Dbf Susceptibility LS].[Coating Degradation Identified], [Dbf Susceptibility LS].[Pipe Wall Degradation], [Dbf Susceptibility LS].[Susceptibility Engineering Judgement], [Dbf Susceptibility LS].[PipeAge10_30 Date], [Dbf Susceptibility LS].[PipeAge30 Date], [Dbf Susceptibility LS].CorrosiveFluid, [Dbf Susceptibility LS].Temp200, [Dbf Susceptibility LS].NoChemicalAdditions, [Dbf Susceptibility LS].[Susceptibility Eng Judgment Value], [Dbf Susceptibility LS].[Susceptibility Eng Judgment Basis] FROM [Dbf Susceptibility LS] WHERE ((([Dbf Susceptibility LS].LineID)=" +
                                _lineid.ToString() + ") AND (([Dbf Susceptibility LS].[Begin Point])<=" +
                                tbBeginPoint.Text + ") AND (([Dbf Susceptibility LS].[End Point])>=" + tbEndPoint.Text +
                                "))";
                            command = new OleDbCommand(queryString, connection);
                            try
                            {
                                reader = command.ExecuteReader();
                                if (reader.Read())
                                {
                                    // Add checked values to the data grid
                                    // Susceptibility Values
                                }
                                reader.Close();
                            }
                            catch (Exception ex)
                            {
                                connection.Close();
                                return -1;
                            }
                        }
                    }
                    //...............
                    // Line Class
                    //...............
                    if (lineclasslookupid > 0)
                    {
                        queryString =
                            "SELECT [Dbf Setup -  Data].SetupID, [Dbf Setup -  Data].ValueData, [Dbf Setup -  Data].Type, [Dbf Setup -  Data].Description, [Dbf Setup -  Data].GUID, [Dbf Setup -  Data].GenericID FROM [Dbf Setup -  Data] WHERE (([Dbf Setup -  Data].SetupID)=" +
                            lineclasslookupid + ") AND (([Dbf Setup -  Data].Type)=\"LineCodeClass\")";
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
                    } // if lineclasslookupid>0          


                    //...............
                    // Line Content
                    if (contentlookupid > 0)
                    {
                        if (connection != null)
                        {
                            queryString =
                                "SELECT [Dbf Setup -  Data].SetupID, [Dbf Setup -  Data].ValueData, [Dbf Setup -  Data].Type, [Dbf Setup -  Data].Description, [Dbf Setup -  Data].GUID, [Dbf Setup -  Data].GenericID FROM [Dbf Setup -  Data] WHERE (([Dbf Setup -  Data].SetupID)=" +
                                contentlookupid + ") AND (([Dbf Setup -  Data].Type)=\"content\")";
                            command = new OleDbCommand(queryString, connection);
                            try
                            {
                                reader = command.ExecuteReader();
                                if (reader.Read())
                                    tbRow.Text = GetFieldDataStr("ValueData", ref reader);
                                reader.Close();
                            }
                            catch (Exception ex)
                            {
                            }
                        } // if connection!=null
                    } // if contentlookupid >0

                    //...............
                    // Set Risk Ranking Colors
                    if (connection != null)
                    {
                        int RRValue = 0;
                        queryString =
                            "SELECT [Dbf RiskRange].LowSus, [Dbf RiskRange].LowMedSus, [Dbf RiskRange].MedSus, [Dbf RiskRange].MedHighSus, [Dbf RiskRange].HighSus, [Dbf RiskRange].LowCon, [Dbf RiskRange].LowMedCon, [Dbf RiskRange].MedCon, [Dbf RiskRange].MedHighCon, [Dbf RiskRange].HighCon, [Dbf RiskRange].LowRisk, [Dbf RiskRange].LowMedRisk, [Dbf RiskRange].MedRisk, [Dbf RiskRange].MedHighRisk, [Dbf RiskRange].HighRisk FROM [Dbf RiskRange]";
                        command = new OleDbCommand(queryString, connection);
                        try
                        {
                            reader = command.ExecuteReader();
                            if (reader.Read())
                            {
                                // Susceptibility
                                RRValue = Convert.ToInt32(tbSusRanking.Text);
                                if (RRValue < GetFieldDataInt("LowSus", ref reader))
                                {
                                    lblSusRanking.BackColor = _rr_low_color;
                                    lblSusRanking.Text = "Low";
                                }
                                else
                                    if (RRValue < GetFieldDataInt("LowMedSus", ref reader))
                                    {
                                        lblSusRanking.BackColor = _rr_low_medium_color;
                                        lblSusRanking.Text = "Low-Medium";
                                    }
                                    else
                                        if (RRValue < GetFieldDataInt("MedSus", ref reader))
                                        {
                                            lblSusRanking.BackColor = _rr_medium_color;
                                            lblSusRanking.Text = "Medium";
                                        }
                                        else
                                            if (RRValue < GetFieldDataInt("MedHighSus", ref reader))
                                            {
                                                lblSusRanking.BackColor = _rr_medium_high_color;
                                                lblSusRanking.Text = "Medium-High";
                                            }
                                            else
                                                if (RRValue < GetFieldDataInt("HighSus", ref reader))
                                                {
                                                    lblSusRanking.BackColor = _rr_high_color;
                                                    lblSusRanking.Text = "High";
                                                }
                                // Consequence
                                RRValue = Convert.ToInt32(tbConRanking.Text);
                                if (RRValue <= GetFieldDataInt("LowCon", ref reader))
                                {
                                    lblConRanking.BackColor = _rr_low_color;
                                    lblConRanking.Text = "Low";
                                }
                                else
                                    if (RRValue <= GetFieldDataInt("LowMedCon", ref reader))
                                    {
                                        lblConRanking.BackColor = _rr_low_medium_color;
                                        lblConRanking.Text = "Low-Medium";
                                    }
                                    else
                                        if (RRValue <= GetFieldDataInt("MedCon", ref reader))
                                        {
                                            lblConRanking.BackColor = _rr_medium_color;
                                            lblConRanking.Text = "Medium";
                                        }
                                        else
                                            if (RRValue <= GetFieldDataInt("MedHighCon", ref reader))
                                            {
                                                lblConRanking.BackColor = _rr_medium_high_color;
                                                lblConRanking.Text = "Medium-High";
                                            }
                                            else
                                                if (RRValue <= GetFieldDataInt("HighCon", ref reader))
                                                {
                                                    lblConRanking.BackColor = _rr_high_color;
                                                    lblConRanking.Text = "High";
                                                }
                                // Overall
                                RRValue = Convert.ToInt32(tbOverallRanking.Text);
                                if (RRValue < GetFieldDataInt("LowRisk", ref reader))
                                {
                                    lblOverallRanking.BackColor = _rr_low_color;
                                    lblOverallRanking.Text = "Low";
                                }
                                else
                                    if (RRValue < GetFieldDataInt("LowMedRisk", ref reader))
                                    {
                                        lblOverallRanking.BackColor = _rr_low_medium_color;
                                        lblOverallRanking.Text = "Low-Medium";
                                    }
                                    else
                                        if (RRValue < GetFieldDataInt("MedRisk", ref reader))
                                        {
                                            lblOverallRanking.BackColor = _rr_medium_color;
                                            lblOverallRanking.Text = "Medium";
                                        }
                                        else
                                            if (RRValue < GetFieldDataInt("MedHighRisk", ref reader))
                                            {
                                                lblOverallRanking.BackColor = _rr_medium_high_color;
                                                lblOverallRanking.Text = "Medium-High";
                                            }
                                            else
                                                if (RRValue < GetFieldDataInt("HighRisk", ref reader))
                                                {
                                                    lblOverallRanking.BackColor = _rr_high_color;
                                                    lblOverallRanking.Text = "High";
                                                }
                            }
                            reader.Close();
                        }
                        catch (Exception ex)
                        {
                        }
                        connection.Close();
                    } // if connection!=null

                } // try...
                finally
                {
//                    connection.Close();
                }

            // Now load the Line Section Risk Ranking Values
            LoadRiskRanking();
            return 0;
        }

       
        //-----------------------------------------------------------------------------------
        // Load the Risk Ranking Values
        //-----------------------------------------------------------------------------------
        private void LoadRiskRanking()
        {
            // Clear Totals
            _suscrisktotal = 0;
            _consrisktotal = 0;
            //........................................
            // Load Susceptibility
            lvSusc.Columns.Clear();
            lvSusc.Columns.Add("Risk Ranking Factor");
            lvSusc.Columns[0].Width = 315;
            lvSusc.Columns.Add("Value", "Value");
            lvSusc.Columns[1].Width = 50;
            lvSusc.Columns[1].TextAlign = HorizontalAlignment.Right;
            lvSusc.View = View.Details;
            _suscrisktotal = 0;
            BPSusRiskRankingValues SusRR = new BPSusRiskRankingValues(_lineid, _beginpoint, _endpoint, _connectionstring);
            int i;
            ListViewItem lvi = null;
            BPRiskRankingItem BBRRitem = new BPRiskRankingItem("",0,"","");
            if (SusRR != null)
            {
                for (i = 0; i <= (SusRR.RrValues.Count - 1); i++)
                {
                    lvi = new ListViewItem();
                    if (SusRR.RrValues.GetItem(i, ref BBRRitem))
                    {
                        lvi.Text = BBRRitem.RiskRankingDesc;
                        lvSusc.Items.Add(lvi);
                        lvi.SubItems.Add(BBRRitem.RiskRankingValue.ToString());
                        _suscrisktotal = _suscrisktotal + BBRRitem.RiskRankingValue;
                    }
                } // for                
            }
            txtSuscTotal.Text = _suscrisktotal.ToString();

            //........................................
            // Load Consequences
            lvCons.Columns.Clear();
            lvCons.Columns.Add("Risk Ranking Factor");
            lvCons.Columns[0].Width = 315;
            lvCons.Columns.Add("Value", "Value");
            lvCons.Columns[1].Width = 50;
            lvCons.Columns[1].TextAlign = HorizontalAlignment.Right;
            lvCons.View = View.Details;
            BPConRiskRankingValues ConRR = new BPConRiskRankingValues(_lineid, _beginpoint, _endpoint, _connectionstring);
            lvi = null;
            _consrisktotal = 0;
            BPRiskRankingItem BPRRitem = new BPRiskRankingItem("", 0, "", "");
            if (ConRR != null)
            {
                for (i = 0; i <= (ConRR.RrValues.Count - 1); i++)
                {
                    lvi = new ListViewItem();
                    if (ConRR.RrValues.GetItem(i, ref BPRRitem))
                    {
                        lvi.Text = BPRRitem.RiskRankingDesc;
                        lvCons.Items.Add(lvi);
                        lvi.SubItems.Add(BPRRitem.RiskRankingValue.ToString());
                        _consrisktotal = _consrisktotal + BPRRitem.RiskRankingValue;
                    }
                } // for                
            }
            txtConsTotal.Text = _consrisktotal.ToString();

            //........................................
            // Load Soil Susceptibility           
            lvSoil.Columns.Clear();
            lvSoil.Columns.Add("Risk Ranking Factor");
            lvSoil.Columns[0].Width = 315;
            lvSoil.Columns.Add("Value", "Value");
            lvSoil.Columns[1].Width = 50;
            lvSoil.Columns[1].TextAlign = HorizontalAlignment.Right;
            lvSoil.View = View.Details;
            double pH = 0;
            double resistivity = 0;
            double Cl = 0;
            double redoxpotential = 0;
            string soiltype = "";
            int RiskPoints = 0;
            string DataDesc = "";
            _soilrisktotal = 0;
            if (GetSoilData(ref pH, ref resistivity, ref Cl, ref redoxpotential, ref soiltype))
            {
                if (GetSoilRiskPoints(pH.ToString(), "pH", ref DataDesc, ref RiskPoints))
                {
                    // pH
                    lvi = new ListViewItem();
                    lvi.Text = "pH: " + DataDesc;
                    lvSoil.Items.Add(lvi);
                    lvi.SubItems.Add(RiskPoints.ToString());
                    _soilrisktotal = _soilrisktotal + RiskPoints;
                }
                // Cl
                if (GetSoilRiskPoints(Cl.ToString(), "Chloride Content", ref DataDesc, ref RiskPoints))
                {
                    lvi = new ListViewItem();
                    lvi.Text = "Chloride Content: " + DataDesc;
                    lvSoil.Items.Add(lvi);
                    lvi.SubItems.Add(RiskPoints.ToString());
                    _soilrisktotal = _soilrisktotal + RiskPoints;
                }
                // Resistivity
                if (GetSoilRiskPoints(resistivity.ToString(), "Soil Resistivity", ref DataDesc, ref RiskPoints))
                {
                    lvi = new ListViewItem();
                    lvi.Text = "Soil Resistivity: " + DataDesc;
                    lvSoil.Items.Add(lvi);
                    lvi.SubItems.Add(RiskPoints.ToString());
                    _soilrisktotal = _soilrisktotal + RiskPoints;
                }
                // Resistivity
                if (GetSoilRiskPoints(redoxpotential.ToString(), "Redox Potential", ref DataDesc, ref RiskPoints))
                {
                    lvi = new ListViewItem();
                    lvi.Text = "Redox Potential: " + DataDesc;
                    lvSoil.Items.Add(lvi);
                    lvi.SubItems.Add(RiskPoints.ToString());
                    _soilrisktotal = _soilrisktotal + RiskPoints;
                }
                // Soil Type
                if (GetSoilDescriptionRiskPoints(soiltype, ref RiskPoints))
                {
                    lvi = new ListViewItem();
                    lvi.Text = "Soil Type: " + DataDesc;
                    lvSoil.Items.Add(lvi);
                    lvi.SubItems.Add(RiskPoints.ToString());
                    _soilrisktotal = _soilrisktotal + RiskPoints;
                }
            }
            txtSoilTotal.Text = _soilrisktotal.ToString();
            // Calculate Totals and Display
            _suscsoilrisktotal = _suscrisktotal + _soilrisktotal;
            txtSuscSoilTotal.Text = _suscsoilrisktotal.ToString();
            // Use BG color to display risk ranking level
            txtSuscSoilTotal.BackColor = lblSusRanking.BackColor;
            txtConsTotal.BackColor = lblConRanking.BackColor;
        }
        


        //-----------------------------------------------------------------------------------
        // Show the Inspection Form
        //-----------------------------------------------------------------------------------
        private void btnInspections_Click(object sender, EventArgs e)
        {
            BPLineSectionInspections lsiform = new BPLineSectionInspections(_zoneid,_connectionstring);
            lsiform.ShowDialog();
            lsiform.Dispose();
        }


        //-----------------------------------------------------------------------------------
        // Show the Line Form
        //-----------------------------------------------------------------------------------
        private void btnShowLine_Click(object sender, EventArgs e)
        {
            if (_allow_open_lines)
            {
                if (_lineid > 0)
                {
                    BPLineDetails ldform = new BPLineDetails(_lineid, _connectionstring, false);
                    ldform.ShowDialog();
                    ldform.Dispose();
                }
            }
        } // LoadFormData



    }
}
