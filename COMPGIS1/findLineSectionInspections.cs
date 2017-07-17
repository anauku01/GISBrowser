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
    public partial class findLineSectionInspections : Form
    {

        private string _connectionstring = "";
        private int _resultscount = 0;
        private int _resultscountgen = 0;
        private int _resultscounteval = 0;
        private bool _selectedevaluations = true;
        private bool _applytocurrentselection = false;

        private int _selectedsetupid = 0;
        private string _selectedmethod = "";
        private string _selectedworkscope = "";

        public string SelectedWorkScope
        {
            set { }
            get { return _selectedworkscope; }
        }

        public bool SelectedEvaluation
        {
            set { }
            get { return _selectedevaluations; }
        }

        public int SelectedSetupID
        {
            set { }
            get { return _selectedsetupid; }
        }

        public string SelectedMethod
        {
            set { }
            get { return _selectedmethod; }
        }

        public bool ApplyToCurrentSelection
        {
            set { }
            get { return _applytocurrentselection; }
        }


        //-----------------------------------------------------------------------------------
        // Constructor
        //-----------------------------------------------------------------------------------
        public findLineSectionInspections(string connectionstring) 
        {
            InitializeComponent();
            _connectionstring = connectionstring;
            if (connectionstring.Length > 0)
            {
                if (LoadFormData())
                {
                    if (cbMethodGeneral.Items.Count > 0)
                    {
                        cbMethodGeneral.SelectedIndex = 0;
                        RunQueryGeneral();
                        SetControls();
                    }
                    if (cbMethodEvals.Items.Count > 0)
                    {
                        cbMethodEvals.SelectedIndex = 0;
                        cbWorkScope.SelectedIndex = 0;
                        RunQueryEval();
                        SetControls();
                    }                    

                }
            }
        }


        //-----------------------------------------------------------------------------------
        // Get string Field Data from the Reader
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
        // Get int Field Data from the Reader
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
        public bool LoadFormData()
        {
            string queryString = "";
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
                try
                {

                    //----------------------
                    // General Tab Setup
                    //----------------------
                    connection.Open();
                    // First Load the Inspection Method Combo
                    queryString = "SELECT [Dbf Setup -  Data].SetupID, [Dbf Setup -  Data].ValueData, [Dbf Setup -  Data].Type, [Dbf Setup -  Data].Description, " +
                                  "[Dbf Setup -  Data].GUID, [Dbf Setup -  Data].GenericID FROM [Dbf Setup -  Data] " +
                                  "WHERE ((([Dbf Setup -  Data].Type)='InspectionMethod')) ORDER BY [Dbf Setup -  Data].ValueData";
                    command = new OleDbCommand(queryString, connection);
                    string MethodName;
                    int MethodID;
                    try
                    {
                        reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            MethodName = GetFieldDataStr("ValueData", ref reader);
                            MethodID = GetFieldDataInt("SetupID", ref reader);
                            if ((MethodName.Length>0) && (MethodID>0)) 
                                cbMethodGeneral.Items.Add(new ListBoxNameValueIntObject { ObjectName = MethodName, ObjectValue = MethodID });
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        connection.Close();
                        return false;
                    }
                    // Leave if the Method combo is not loaded
                    if (cbMethodGeneral.Items.Count == 0)
                        return false;


                    //----------------------
                    // Evaluation Tab Setup
                    //----------------------
                    // Load the Work Scopes Combo
                    cbWorkScope.Items.Clear();
                    cbWorkScope.Items.Add("All");
                    queryString = "SELECT [Dbf Setup - Workscope].PerId, [Dbf Setup - Workscope].[Work Scope] FROM [Dbf Setup - Workscope]";
                    command = new OleDbCommand(queryString, connection);
                    try
                    {
                        reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            MethodName = GetFieldDataStr("Work Scope", ref reader);
                            if (MethodName.Length > 0)
                                cbWorkScope.Items.Add(MethodName);
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        connection.Close();
                        return false;
                    }

                } // try...

                finally
                {
                    connection.Close();
                }
            return true;
        } // LoadFormData



        //-----------------------------------------------------------------------------------
        // Run the Query for Inspections of the selected Method
        //-----------------------------------------------------------------------------------
        public int LoadInspMethod(int SetupID)
        {
            int count = 0;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string cmdStr = "SELECT COUNT(*) FROM [Dbf Inspection Range] " +
                                  "INNER JOIN [Dbf Line Section] ON [Dbf Inspection Range].LineID = [Dbf Line Section].LineID WHERE ((([Dbf Line Section].[Begin Point])>=[Dbf Inspection Range].[Begin Point]) AND (([Dbf Line Section].[End Point])<=[Dbf Inspection Range].[End Point]) AND (([Dbf Inspection Range].Method)=" + SetupID + "))";
                    //string cmdStr = "SELECT [Dbf Line Section].[Zone ID], [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point], " +
                    //              "[Dbf Inspection Range].Method FROM [Dbf Inspection Range] " +
                    //              "INNER JOIN [Dbf Line Section] ON [Dbf Inspection Range].LineID = [Dbf Line Section].LineID " +
                    //              "WHERE ((([Dbf Line Section].[Begin Point])>=[Dbf Inspection Range].[Begin Point]) AND (([Dbf Line Section].[End Point])<=[Dbf Inspection Range].[End Point]) AND (([Dbf Inspection Range].Method)=" + SetupID + "))";

                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    count = (Int32)cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
            }
            return count;
        } // Run the Query with method



        //-----------------------------------------------------------------------------------
        // Run the Query for Inspections of the selected Method
        //-----------------------------------------------------------------------------------
        public int LoadEvalMethod(string MethodStr, string WorkScope)
        {
            Int32 count = 0;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
                connection.Open();
                try
                {
                    string whereclause = "";
                    string cmdStr =
                        "SELECT COUNT(*) FROM [Dbf Inspections]";
//                    "SELECT COUNT(*) [Dbf Inspections].[Zone ID], [Dbf Inspections].[WS Name] FROM [Dbf Inspections]";

                    if (string.Compare(MethodStr, "All", true) != 0)
                    {
                        whereclause = "(Method='" + MethodStr + "')";
                    }
                    if (string.Compare(WorkScope, "All", true) != 0)
                    {
                        if (whereclause.Length > 0)
                            whereclause += " OR ";
                        whereclause = " ([Dbf Inspections].[WS Name] ='" + WorkScope + "')";
                    }
                    // Build the Final Query String
                    if (whereclause.Length > 0)
                    {
                        cmdStr+= " WHERE " + whereclause;
                    }

                    OleDbCommand cmd = new OleDbCommand(cmdStr, connection);
                    count =  (Int32) cmd.ExecuteScalar();
                    if (count == null) return 0;
                }
                catch (Exception ex)
                {
                    string m = ex.Message;
                }
                finally
                {
                    connection.Close();
                }
            }
            return count;
        } // Run the Query with method


        //-----------------------------------------------------------------------------------
        // Tab Changed - Set Form Values
        //-----------------------------------------------------------------------------------
        private void SetTabChange()
        {
            switch (tcMain.SelectedIndex)
            {
                case (0):
                    {
                        lblResults.Text = _resultscountgen.ToString();
                        _resultscount = _resultscountgen;
                        break;
                    }
                case (1):
                    {
                        lblResults.Text = _resultscounteval.ToString();
                        _resultscount = _resultscounteval;
                        break;
                    }
            }
        }


        //-----------------------------------------------------------------------------------
        // General Method was changed
        //-----------------------------------------------------------------------------------
        private void RunQueryGeneral()
        {
            if (cbMethodGeneral.SelectedIndex < 0) return;
            int SetupID = ((ListBoxNameValueIntObject)(cbMethodGeneral.Items[cbMethodGeneral.SelectedIndex])).ObjectValue;
            int count = LoadInspMethod(SetupID);
            _resultscount = count;
//            _resultscountgen = count;
//            lblResults.Text = _resultscount.ToString();
            SetControls();
        }


        //-----------------------------------------------------------------------------------
        // Eval Method was changed
        //-----------------------------------------------------------------------------------
        private void RunQueryEval()
        {
            if (cbMethodGeneral.SelectedIndex < 0) return;
            if (cbWorkScope.SelectedIndex < 0) return;
            int MethodIndex = cbMethodGeneral.SelectedIndex;
            string WorkScope = cbWorkScope.Items[cbWorkScope.SelectedIndex].ToString();
            int count = LoadEvalMethod(cbMethodEvals.Items[cbMethodEvals.SelectedIndex].ToString(), WorkScope);
            _resultscount = count;
            _resultscounteval = count;
            lblResults.Text = _resultscount.ToString();
            SetControls();
        }

        //-----------------------------------------------------------------------------------
        // Set Controls based upon 
        //-----------------------------------------------------------------------------------
        private void SetControls()
        {
            btnOK.Enabled = (_resultscount > 0);
        } 

        //-----------------------------------------------------------------------------------
        // Method was changed
        //-----------------------------------------------------------------------------------
        private void cbMethod_SelectedIndexChanged(object sender, EventArgs e)
        {
            RunQueryGeneral();
        }

        //-----------------------------------------------------------------------------------
        // OK Was clicked, set the selectedsetupid
        //-----------------------------------------------------------------------------------
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (tcMain.SelectedIndex == 0)
            {
                if (cbMethodGeneral.SelectedIndex >= 0)
                {
                    _selectedevaluations = false;
                    _selectedsetupid = ((ListBoxNameValueIntObject)(cbMethodGeneral.Items[cbMethodGeneral.SelectedIndex])).ObjectValue;
                    _selectedmethod = ((ListBoxNameValueIntObject)(cbMethodGeneral.Items[cbMethodGeneral.SelectedIndex])).ObjectName;
                }               
            }
            else
            {
                if ((cbMethodEvals.SelectedIndex >= 0) && (cbWorkScope.SelectedIndex >= 0))
                {
                    _selectedevaluations = true;
                    _selectedworkscope = cbWorkScope.Items[cbWorkScope.SelectedIndex].ToString();
                    _selectedmethod = cbMethodEvals.Items[cbMethodEvals.SelectedIndex].ToString();
                }               
                
            }
        }

        //-----------------------------------------------------------------------------------
        // Tab Changed
        //-----------------------------------------------------------------------------------
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetTabChange();
            SetControls();
        }


        //-----------------------------------------------------------------------------------
        // Eval Method Selected
        //-----------------------------------------------------------------------------------
        private void cbMethodEvals_SelectedIndexChanged(object sender, EventArgs e)
        {
//            RunQueryEval();
//            SetControls();
        }

        private void cbMethodEvals_SelectionChangeCommitted(object sender, EventArgs e)
        {
            RunQueryEval();
            SetControls();
        }

        private void cbApplyToSelection_CheckedChanged(object sender, EventArgs e)
        {
            _applytocurrentselection = cbApplyToSelection.Checked;
        } 

    }
}
