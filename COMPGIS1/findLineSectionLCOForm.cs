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

    public partial class findLineSectionLCOForm : Form
    {

        private string _connectionstring = "";
        private int _default_LCO = 6;
        private int _lco = 0;
        private string _lcostr = "All";
        private bool _applytocurrentselection = false;
        private bool _inclusivebelow = false;

        public string LCODesc
        {
            set { }
            get { return _lcostr; }
        }

        public int LCO {
            set { } 
            get { return _lco; }
        }

        public bool ApplyToCurrentSelection
        {
            set { }
            get { return _applytocurrentselection; }
        }

        public bool InclusiveBelow
        {
            set { }
            get { return _inclusivebelow; }
        }

        public findLineSectionLCOForm(string connectionstring)
        {
            InitializeComponent();
            {
                _connectionstring = connectionstring;
                LoadFormData();
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
        public void LoadFormData()
        {
            cbLCO.SelectedIndex = _default_LCO;
            RunQueryForLCO(_default_LCO);
            cbLCO.Focus();
        } // LoadFormData


        //-----------------------------------------------------------------------------------
        // Show number of records meeting the criteria
        //-----------------------------------------------------------------------------------
        private void RunQueryForLCO(int lco)
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
                    // Build the query to search with
                    string LCOstr = "";
                    if (lco > 5) // Do any with LCO
                        LCOstr = "> 0";
                    else
                        LCOstr = "= " + LCO.ToString();
                    queryString =
                        "SELECT [Dbf Line Section].[Zone ID], [Dbf Consequence].LCO, [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point], " +
                        "[Dbf Line Section].[Line Section], [Dbf Line Section].MapFeatureID " +
                        "FROM [Dbf Consequence] INNER JOIN [Dbf Line Section] ON [Dbf Consequence].LineID = [Dbf Line Section].LineID " +
                        "WHERE ((([Dbf Consequence].LCO) " + LCOstr + ") AND (([Dbf Line Section].[Begin Point])>=[Dbf Consequence].[Begin Point]) AND (([Dbf Line Section].[End Point])<=[Dbf Consequence].[End Point]))";
                    command = new OleDbCommand(queryString, connection);
                    int count = (Int32) command.ExecuteScalar();

                }
                catch (Exception ex)
                {
                    connection.Close();
                }
            } 
        }

        //-----------------------------------------------------------------------------------
        // User Clicked OK, set LCO
        //-----------------------------------------------------------------------------------
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (cbLCO.SelectedIndex > 5)
                _lco = -1;
            else
                _lco = cbLCO.SelectedIndex;
            if (_inclusivebelow)
            {
                _lcostr = "<= " +cbLCO.Items[cbLCO.SelectedIndex].ToString();                
            }
            else
            {
                _lcostr = cbLCO.Items[cbLCO.SelectedIndex].ToString();                
            }
        }

        //-----------------------------------------------------------------------------------
        // User Changed LCO to search for
        //-----------------------------------------------------------------------------------
        private void cbLCO_SelectedIndexChanged(object sender, EventArgs e)
        {
            RunQueryForLCO(cbLCO.SelectedIndex);
            SetControls();
        }

        //-----------------------------------------------------------------------------------
        // User Changed Apply to Selection Checkbox
        //-----------------------------------------------------------------------------------
        private void cbApplyToSelection_CheckedChanged(object sender, EventArgs e)
        {
            _applytocurrentselection = cbApplyToSelection.Checked;
        }

        //-----------------------------------------------------------------------------------
        // User Changed Inclusive Checkbox
        //-----------------------------------------------------------------------------------
        private void cbInclusive_CheckedChanged(object sender, EventArgs e)
        {
            _inclusivebelow = cbInclusive.Checked;
        }

        //-----------------------------------------------------------------------------------
        // Set controls based upon user selection
        //-----------------------------------------------------------------------------------
        private void SetControls()
        {
            // dropdown options...
            //0 None
            //1: < 72 Hours
            //2: 72 Hours
            //3: 7 Days
            //4: Days
            //5: 30 Days
            //6: Any LCO
            cbInclusive.Enabled = ((cbLCO.SelectedIndex > 1) && (!(cbLCO.SelectedIndex >5)));
            if (!cbInclusive.Enabled)
            {
                _inclusivebelow = false;
            }                
            cbApplyToSelection.Enabled = (cbLCO.SelectedIndex > 0);            
        }

    }
}
