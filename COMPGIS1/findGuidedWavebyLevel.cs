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
    public partial class findGuidedWavebyLevel : Form
    {
        private string _connectionstring = "";
        private int _selectedlevel = 0;
        public int SelectedLevel
        {
            set { }
            get { return _selectedlevel; }
        }

        private bool _applytocurrentselection = false;

        public bool ApplyToCurrentSelection
        {
            set { }
            get { return _applytocurrentselection; }
        }



        //-----------------------------------------------------------------------------------
        // Constructor
        //-----------------------------------------------------------------------------------
        public findGuidedWavebyLevel(string connectionstring)
        {
            InitializeComponent();
            _connectionstring = connectionstring;
            if (_connectionstring.Length > 0)
            {                
                LoadFormData();
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
                {
                    string tststr = rdr[idx].ToString();
                    if (tststr.Length>0)
                        retval = Convert.ToInt32(tststr);
                }                    
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
                    connection.Open();
                    queryString ="SELECT DISTINCT [Dbf Inspections Guided Wave].ResultsLevel FROM [Dbf Inspections Guided Wave]";
                    command = new OleDbCommand(queryString, connection);
                    int RL = 0;

                    try
                    {
                        reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            RL = GetFieldDataInt("ResultsLevel", ref reader);
                            if (RL > 0 )
                                cbLevel.Items.Add(new ListBoxNameValueIntObject { ObjectName = RL.ToString(), ObjectValue = RL });
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        connection.Close();
                        return false;
                    }

                    if (cbLevel.Items.Count == 0)
                    {
                        btnOK.Enabled = false;
                        lblErrorMsg.Visible = true;
                        lblErrorMsg.Text = "No Guided Wave Inspections results available.";
                    }

                } // try...
                finally
                {
                    connection.Close();
                }
            return true;
        } // LoadFormData

        //-----------------------------------------------------------------------------------
        // Set Form Controls based upon selections
        //-----------------------------------------------------------------------------------
        public void SetControls()
        {
            btnOK.Enabled = (cbLevel.SelectedIndex >= 0);
        }


        //-----------------------------------------------------------------------------------
        // User Selected a Level
        //-----------------------------------------------------------------------------------
        private void cbLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetControls();
        } 

        //-----------------------------------------------------------------------------------
        // User Clicked OK
        //-----------------------------------------------------------------------------------
        private void btnOK_Click(object sender, EventArgs e)
        {
            _selectedlevel = ((ListBoxNameValueIntObject)(cbLevel.Items[cbLevel.SelectedIndex])).ObjectValue;
        }

        //-----------------------------------------------------------------------------------
        // User Changed Apply to Selection Checkbox
        //-----------------------------------------------------------------------------------
        private void cbApplyToSelection_CheckedChanged(object sender, EventArgs e)
        {
            _applytocurrentselection = cbApplyToSelection.Checked;
        }


    }
}
