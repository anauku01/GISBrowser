using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace COMPGIS1
{
    public partial class findLineSectionsRiskRanking : Form
    {

        private string _connectionstring = "";
        private BPSusRiskRankingValues _rrsusc;
        private BPConRiskRankingValues _rrcons;
        private bool _applytocurrentselection = false;
        private string _sqlstatementcons = "";
        private string _sqlstatementsusc = "";
        private string _selectdescription = "";

        public string SelectDescription
        {
            get { return _selectdescription; }
            set { }
        }

        public string SQLStatementSusc
        {
            get { return _sqlstatementsusc; }
            set { }
        }

        public string SQLStatementCons
        {
            get { return _sqlstatementcons; }
            set { }
        }

        public bool ApplyToCurrentSelection
        {
            set { }
            get { return _applytocurrentselection; }
        }

        public int SelectCount;



        // Constructor
        public findLineSectionsRiskRanking(string connectionstring)
        {
            InitializeComponent();
            clbSusc.DisplayMember = "ObjectName";
            clbSusc.ValueMember = "ObjectValue";
            clbCons.DisplayMember = "ObjectName";
            clbCons.ValueMember = "ObjectValue";
            _connectionstring = connectionstring;
            _rrsusc = new BPSusRiskRankingValues(_connectionstring);
            _rrcons = new BPConRiskRankingValues(_connectionstring);
            if ((_rrsusc != null) && (_rrcons != null))
            {
                LoadFormData();
            }
        }

        //.....................................
        // Load Form Data
        //.....................................
        private void LoadFormData()
        {
            // Susceptibility
            clbSusc.Items.Clear();
            int i;
            for (i = 0; i < (_rrsusc.RrValues.Count - 1); i++)
            {
                BPRiskRankingItem RRItem = new BPRiskRankingItem("",0,"","");
                if (_rrsusc.RrValues.GetItem(i, ref RRItem))
                {
                    clbSusc.Items.Add(RRItem.RiskRankingDesc);
                }
            }
            // Consequence
            clbCons.Items.Clear();
            for (i = 0; i < (_rrcons.RrValues.Count - 1); i++)
            {
                BPRiskRankingItem RRItem = new BPRiskRankingItem("", 0, "", "");
                if (_rrcons.RrValues.GetItem(i, ref RRItem))
                {
                    //clbCons.Items.Add(RRItem.RiskRankingDesc);
                    clbCons.Items.Add(new ListBoxNameValueStringObject { ObjectName = RRItem.RiskRankingDesc, ObjectValue = RRItem.RiskRankingFieldNameLS});

                }
            }
        }

        // Build SQL Statement if the selections are valid
        private bool BuildSQLStatement()
        {
            int count = 0;
            string logicalcondition = " OR ";
            if (rbAND.Checked)
                logicalcondition = " AND ";
            _selectdescription = "Factors selected:\n";
            // Build the Susc SQL Statement First
            if (clbSusc.CheckedItems.Count > 0)
            {
                BPRiskRankingItem BPRRItem = new BPRiskRankingItem("", 0, "", "");
                string SQLStatement = "SELECT [Dbf Line Section].[Zone ID], [Dbf Line Section].[Line Section], [Dbf Line Section].MapFeatureID, [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point], [Dbf Susceptibility LS].Under_Road, [Dbf Susceptibility LS].Under_Bldg, [Dbf Susceptibility LS].Under_RR, " +
                                      "[Dbf Susceptibility LS].Under_Tower_Footer, [Dbf Susceptibility LS].Under_Trans_Line, [Dbf Susceptibility LS].Over_Under_River, [Dbf Susceptibility LS].In_To_Out_Of_Bldg, [Dbf Susceptibility LS].In_Wall, " +
                                      "[Dbf Susceptibility LS].Underground_Tee, [Dbf Susceptibility LS].Replaced, [Dbf Susceptibility LS].Mtl_Chg, [Dbf Susceptibility LS].Inspected, [Dbf Susceptibility LS].Above_Groundwater_Level, " +
                                      "[Dbf Susceptibility LS].Earthen_Fill_Material_Change, [Dbf Susceptibility LS].Pipe, [Dbf Susceptibility LS].PipeAge, [Dbf Susceptibility LS].Leak_History, [Dbf Susceptibility LS].NotCoated, [Dbf Susceptibility LS].SoilOutOfSpec, " +
                                      "[Dbf Susceptibility LS].GroundSettlement, [Dbf Susceptibility LS].PipesInArea, [Dbf Susceptibility LS].SteamLine, [Dbf Susceptibility LS].WallThinner, [Dbf Susceptibility LS].Potential, [Dbf Susceptibility LS].Uncorrected, " +
                                      "[Dbf Susceptibility LS].BackfillUnacceptable, [Dbf Susceptibility LS].RectifierOperational, [Dbf Susceptibility LS].[Internal Erosion Corrosion], [Dbf Susceptibility LS].[Soil Characteristics Unknown], " +
                                      "[Dbf Susceptibility LS].[Recorded Transient Not Corrected], [Dbf Susceptibility LS].[Within 10 ft Transmission Line Footer], [Dbf Susceptibility LS].[No Coating Inspection Performed], " +
                                      "[Dbf Susceptibility LS].[Coating Degradation Identified], [Dbf Susceptibility LS].[Pipe Wall Degradation], [Dbf Susceptibility LS].[Susceptibility Engineering Judgement], [Dbf Susceptibility LS].[PipeAge10_30 Date], " +
                                      "[Dbf Susceptibility LS].[PipeAge30 Date], [Dbf Susceptibility LS].CorrosiveFluid, [Dbf Susceptibility LS].Temp200, [Dbf Susceptibility LS].NoChemicalAdditions, [Dbf Susceptibility LS].[Susceptibility Eng Judgment Value], " +
                                      "[Dbf Susceptibility LS].[Susceptibility Eng Judgment Basis] " +
                                      "FROM [Dbf Susceptibility LS] INNER JOIN [Dbf Line Section] ON [Dbf Susceptibility LS].LineID = [Dbf Line Section].LineID " + "WHERE ((([Dbf Line Section].[Begin Point])<=[Dbf Susceptibility LS].[Begin Point]) AND (([Dbf Line Section].[End Point])>=[Dbf Susceptibility LS].[End Point])) AND (";
                //            "FROM [Dbf Susceptibility LS] INNER JOIN [Dbf Line Section] ON [Dbf Susceptibility LS].LineID = [Dbf Line Section].LineID " + "WHERE ((([Dbf Line Section].[Begin Point])<=[Dbf Susceptibility LS].[Begin Point]) AND (([Dbf Line Section].[End Point])>=[Dbf Susceptibility LS].[End Point]) ";
                bool firstcond = true;
                foreach (int indexChecked in clbSusc.CheckedIndices)
                {
                    if (_rrsusc.RrValues.GetItem(indexChecked, ref BPRRItem))
                    {
                        if ((string.Compare(BPRRItem.RiskRankingFieldNameLS, "Pipe") == 0) || (string.Compare(BPRRItem.RiskRankingFieldNameLS, "PipeAge") == 0) || (string.Compare(BPRRItem.RiskRankingFieldNameLS, "Leak_History") == 0))
                        {
                            if (!firstcond)
                                SQLStatement += string.Format(logicalcondition + "(([Dbf Susceptibility LS].[{0}])={1})", BPRRItem.RiskRankingFieldNameLS, BPRRItem.RiskRankingValue);
                            else
                                SQLStatement += string.Format("(([Dbf Susceptibility LS].[{0}])={1})", BPRRItem.RiskRankingFieldNameLS, BPRRItem.RiskRankingValue);
                            _selectdescription += string.Format("\n{0}",BPRRItem.RiskRankingDesc);
                        }
                        else
                        {
                            if (!firstcond)
                                SQLStatement += string.Format(logicalcondition + "(([Dbf Susceptibility LS].[{0}])<>0)", BPRRItem.RiskRankingFieldNameLS);
                            else
                                SQLStatement += string.Format("(([Dbf Susceptibility LS].[{0}])<>0)", BPRRItem.RiskRankingFieldNameLS);
                            _selectdescription += string.Format("\n{0}",BPRRItem.RiskRankingDesc);
                        }
                        count++;
                        firstcond = false;
                    }
                }
                if (count > 0)
                    _sqlstatementsusc = SQLStatement += ")";
                else
                    _sqlstatementsusc = "";
            } // if Susc count > 0



            // Build the Cons SQL Statement First
            if (clbCons.CheckedItems.Count > 0)
            {
                BPRiskRankingItem BPRRItem = new BPRiskRankingItem("", 0, "", "");
                string SQLStatement = "SELECT [Dbf Line Section].[Zone ID], [Dbf Consequence].LineID, [Dbf Line Section].[Line Section], [Dbf Line Section].MapFeatureID, [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point], [Dbf Consequence].[Collateral Damage], [Dbf Consequence].[Loss of Generation], " +
                                      "[Dbf Consequence].[Outage Duration], [Dbf Consequence].Remarks, [Dbf Consequence].[Cost of Repair], [Dbf Consequence].[If Break], [Dbf Consequence].[If Leak], [Dbf Consequence].[Will Failure Affect Worker Safety], " +
                                      "[Dbf Consequence].GUID, [Dbf Consequence].SafetyRelated, [Dbf Consequence].Radiological, [Dbf Consequence].EPA, [Dbf Consequence].LCO, [Dbf Consequence].Licensing, [Dbf Consequence].Repair1Million, " +
                                      "[Dbf Consequence].Leak_Consequences, [Dbf Consequence].Break_Consequences, [Dbf Consequence].Occlusion_Consequences, [Dbf Consequence].Failure_Worker_Safety, [Dbf Consequence].Failure_Radioactive_Ground_Cont, " +
                                      "[Dbf Consequence].[Failure_Non-Radioactive_Cont], [Dbf Consequence].Failure_Airborne_Cont, [Dbf Consequence].Failure_Loss_Generation, [Dbf Consequence].Failure_Operator_Work_Around, " +
                                      "[Dbf Consequence].Failure_Collateral_Damage, [Dbf Consequence].Failure_Affect_Nuclear_Safety, [Dbf Consequence].[Impact Assesment], [Dbf Consequence].[Corrosion Risk], [Dbf Consequence].[Pipe Break], " +
                                      "[Dbf Consequence].[Break Origin], [Dbf Consequence].[Break Status], [Dbf Consequence].[Cause of Break], [Dbf Consequence].[Date Break was Reported], [Dbf Consequence].[Leaks from Similar Pipes in Area], " +
                                      "[Dbf Consequence].[Suspected Degradation], [Dbf Consequence].[Tubercules Height 10 Dia], [Dbf Consequence].[Risk Adjusted Weight (PRA)], [Dbf Consequence].Line_Service_Unit, [Dbf Consequence].FailureHistory, " +
                                      "[Dbf Consequence].FlowLossOcclusion, [Dbf Consequence].LeakDetectionSystem, [Dbf Consequence].LeakMadeUp, [Dbf Consequence].NRC, [Dbf Consequence].UnderBldg2, [Dbf Consequence].tritium, [Dbf Consequence].RadWaste, " +
                                      "[Dbf Consequence].[Consequence Engineering Judgement], [Dbf Consequence].NSIAC, [Dbf Consequence].DEQ, [Dbf Consequence].[Consequence Eng Judgment Value], [Dbf Consequence].[Consequence Eng Judgment Basis] " +
                                      "FROM [Dbf Consequence] INNER JOIN [Dbf Line Section] ON [Dbf Consequence].LineID = [Dbf Line Section].LineID " +
                                      "WHERE ((([Dbf Line Section].[Begin Point])>=[Dbf Consequence].[Begin Point]) AND (([Dbf Line Section].[End Point])<=[Dbf Consequence].[End Point])) AND (";
                bool firstcond = true;
                foreach (int indexChecked in clbCons.CheckedIndices)
                {
                    if (_rrcons.RrValues.GetItem(indexChecked, ref BPRRItem))
                    {
                        if (string.Compare(BPRRItem.RiskRankingFieldNameLS, "LCO") == 0)
                        {
                            if (!firstcond)
                                SQLStatement += string.Format(logicalcondition + "(([Dbf Consequence].[{0}])={1})", BPRRItem.RiskRankingFieldNameLS, BPRRItem.RiskRankingValue);
                            else
                                SQLStatement += string.Format("(([Dbf Consequence].[{0}])={1})", BPRRItem.RiskRankingFieldNameLS, BPRRItem.RiskRankingValue);
                            _selectdescription += string.Format("\n{0}",BPRRItem.RiskRankingDesc);
                        }
                        else
                        {
                            if (!firstcond)
                                SQLStatement += string.Format(logicalcondition + "(([Dbf Consequence].[{0}])<>0)", BPRRItem.RiskRankingFieldNameLS);
                            else
                                SQLStatement += string.Format("(([Dbf Consequence].[{0}])<>0)", BPRRItem.RiskRankingFieldNameLS);
                            _selectdescription += string.Format("\n{0}",BPRRItem.RiskRankingDesc);
                        }
                        count++;
                        firstcond = false;
                    }
                }
                if (count > 0)
                    _sqlstatementcons = SQLStatement += ")";
                else
                    _sqlstatementcons = "";
            } // if Susc count > 0

            if (count>1)
                _selectdescription += string.Format("\n\nLogical condition: '{0}'", logicalcondition);

            return (count > 0);
        }


        // Build SQL Statement if the selections are valid
        private void findLineSectionsRiskRanking_FormClosing(object sender, FormClosingEventArgs e)
        {
            if ((this.DialogResult == DialogResult.OK) && ((clbSusc.CheckedItems.Count > 0) || (clbCons.CheckedItems.Count > 0)))
            {
                e.Cancel = (!BuildSQLStatement());
            }             
        }

        private void cbApplyToSelection_CheckedChanged(object sender, EventArgs e)
        {
            _applytocurrentselection = cbApplyToSelection.Checked;
        }


    }
}
