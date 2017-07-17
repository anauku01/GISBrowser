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
    public partial class SystemSelectionForm : Form
    {

        private string _connectionstring = "";
        private int _defaultunit = 0;
        private bool _loadingsystems = false;

        private int _systemid = -1;
        private string _systemdesc = "";

        public int SystemID
        {
            get
            {
                return this._systemid;
            }
            set { } 
        }
    
        public string SystemDesc
        {
            get
            {
                return this._systemdesc;
            }
            set { } 
        }


        //-----------------------------------------------------------------------------------
        // Contructor
        //-----------------------------------------------------------------------------------
        public SystemSelectionForm(int defaultunit,  string connectionstring)
        {
            InitializeComponent();
            _connectionstring = connectionstring;
            _defaultunit = defaultunit;
            if (_connectionstring.Length > 0)
            {
                LoadFormData(defaultunit);
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
        public int LoadFormData(int defaultunit)
        {
            int contentlookupid = -1;
            int lineclasslookupid = -1;
            string queryString = "";
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);

            cbUnits.Items.Clear();
            if (connection != null)
            {
                connection.Open();
                queryString = "SELECT DISTINCT [Dbf Systems].Unit FROM [Dbf Systems]";

                command = new OleDbCommand(queryString, connection);
                try
                {
                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        cbUnits.Items.Add(GetFieldDataStr("Unit", ref reader));
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

            int idx = cbUnits.Items.IndexOf(defaultunit.ToString());
            if ((idx > 0) && (idx <= cbUnits.Items.Count - 1))
            {
                cbUnits.SelectedIndex = idx;
            }
        // DONE
            return 0;
            SetControls();
        } // loadformdata 



        //-----------------------------------------------------------------------------------
        // Load the Systems List Box when the Unit changes
        //-----------------------------------------------------------------------------------
        private void LoadListBox(string unitnumber)
        {
            if (_loadingsystems) return;
            _loadingsystems = true;
            string queryString = "";

            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);
            try
            {
                lbSystems.Items.Clear();
                lbSystems.DisplayMember = "ObjectName";
                lbSystems.ValueMember = "ObjectValue";

                connection = new OleDbConnection(_connectionstring);
                if (connection != null)
                {
                    connection.Open();
                    queryString = "SELECT [Dbf Systems].Unit, [Dbf Systems].System, [Dbf Systems].SystemID FROM [Dbf Systems] WHERE ((([Dbf Systems].Unit)='" + unitnumber.ToString() + "'))";
                    command = new OleDbCommand(queryString, connection);
                    try
                    {
                        reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            string Str = GetFieldDataStr("System", ref reader);
                            int SID = GetFieldDataInt("SystemID", ref reader);
                            if ((Str.Length>0) & (SID>0))
                                lbSystems.Items.Add(new ListBoxNameValueIntObject { ObjectName = Str, ObjectValue =  SID });
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        string sss = ex.Message.ToString();
                    }
                    connection.Close();
                    if (lbSystems.Items.Count > 0)
                        lbSystems.SelectedIndex = 0;
                } // if connection!=null
            }
            finally
            {
                _loadingsystems = false;
            }

        }

        //-----------------------------------------------------------------------------------
        // Units combo has changed
        //-----------------------------------------------------------------------------------
        private void cbUnits_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbUnits.SelectedIndex >= 0)
                LoadListBox(cbUnits.Items[cbUnits.SelectedIndex].ToString());
            SetControls();
        }


        //-----------------------------------------------------------------------------------
        // Set the controls
        //-----------------------------------------------------------------------------------
        private void SetControls()
        {
            btnOK.Enabled = (lbSystems.SelectedIndex >= 0);
        }

        //-----------------------------------------------------------------------------------
        // OK button clicked
        //-----------------------------------------------------------------------------------
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (lbSystems.SelectedIndex >= 0)
            {
                _systemid = ((ListBoxNameValueIntObject)(lbSystems.Items[lbSystems.SelectedIndex])).ObjectValue;
                if (_systemid > 0)
                    _systemdesc = ((ListBoxNameValueIntObject)(lbSystems.Items[lbSystems.SelectedIndex])).ObjectName;                
                else
                    _systemid = -1;
            }
        }

        //-----------------------------------------------------------------------------------
        // OK button clicked
        //-----------------------------------------------------------------------------------
        private void lbSystems_DoubleClick(object sender, EventArgs e)
        {
            if (lbSystems.SelectedIndex >= 0)
                btnOK.PerformClick();
        }



    }
}
