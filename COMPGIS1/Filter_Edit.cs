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
    public partial class Filter_Edit : Form
    {

        private string _connectionstring;
        private bool _editmode = false;
        private LayerSystemUse _layerindex;
        private  LayerDefs _layers = null;
        private string _fieldsourcequery;
        private SQLFilterClass _originalfilter = null;
        private List<string> _fieldtypes;

        private int _normalformheight = 520;
        private int _normalscheight = 430;
        private int _editpanelheight = 100;
        private bool _savemode;
        private bool SaveMode
        {
            get
            {
                return this._savemode;
            }
            set 
            { 
                this._savemode = SetSaveMode(value); 
            } 
        }
    

        // Layer related constants/vars
        private const int su_layerLines = 1;
        private const int su_layerLineSections = 2;
        private const int su_layerBoringLocations = 3;
        private const int su_layerCorrosionProbes = 4;

        // Public information to return to caller
        public SQLFilterClass SQLFilter{set;get;}

        // -----------------------------------------------------------------------
        // Constructor - Overload 1
        // layerindex indicates the layer database to be filtered
        // -----------------------------------------------------------------------
        public Filter_Edit(LayerSystemUse layerindex, LayerDefs layers, string connectionstring, bool savemode)
        {
            InitializeComponent();
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            try
            {
                SaveMode = savemode;
                _editmode = false;
                _layerindex = layerindex;
                _connectionstring = connectionstring;
                SQLFilter = new SQLFilterClass();
                _fieldtypes = new List<string>();
                _layers = layers;
                // Set up the form according to the layer being filtered
                SetupFormData();
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
        }

        // -----------------------------------------------------------------------
        // Constructor - Overload 2
        // layerindex indicates the layer database to be filtered
        // -----------------------------------------------------------------------
        public Filter_Edit(SQLFilterClass filter, LayerDefs layers, string connectionstring, bool savemode)
        {
            InitializeComponent();
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            try
            {
                SaveMode = savemode;
                _editmode = true;
                if (filter != null)
                {
                    _layerindex = filter.LayerIndex;
                    _connectionstring = connectionstring;
                    _originalfilter = filter;
                    _layers = layers;
                    _fieldtypes = new List<string>();
                    SQLFilter = new SQLFilterClass(_originalfilter);
                    // Set up the form according to the layer being filtered
                    if (SQLFilter.ConditionCount == _originalfilter.ConditionCount)
                    {

                    }
                    SetupFormData();
                }
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
        }

        // -----------------------------------------------------------------------
        // Form Load Event
        // -----------------------------------------------------------------------
        private void Filter_Edit_Load(object sender, EventArgs e)
        {
            if (_savemode)
            {
                tbFilterName.Select();
            }
        }


        // -----------------------------------------------------------------------
        // Set up the form
        // -----------------------------------------------------------------------
        private bool SetupFormData()
        {
            // Set Conditions ComboBox
            cbCondition.DisplayMember = "ObjectName";
            cbCondition.ValueMember = "ObjectValue";
            _fieldsourcequery = _layers.BaseFilterQuery(_layerindex);

            // Load the source fields
            LoadSourceFieldList();
            if (lbFieldName.Items.Count > 0)
            {
                lbFieldName.SelectedIndex = 0;
                LoadValuesItems();
            }
            // Load any conditions that may exist
            LoadConditionList();
            tbFilterName.Text = SQLFilter.FilterName;
            tbFilterDescription.Text = SQLFilter.FilterDescription;
            cbCondition.SelectedIndex=0;
            lbFieldName.Focus();
            SetControls();
            return true;
        }

        // -----------------------------------------------------------------------
        //  Set the form up for SaveMode value
        //  In Savemode (default) the filter Name and Desc fields are shown 
        //  When SaveMode is false, the filter Name and Desc fields are hidden
        //  and the form height is adjusted (smaller)
        // -----------------------------------------------------------------------
        private bool SetSaveMode(bool savemode)
        {
            scMain.Panel1Collapsed = (! savemode);
            if (savemode) 
            {
                this.Size = new Size(this.Size.Width,_normalformheight);
                scMain.Height = _normalscheight;
            }
            else
            {
                scMain.Height = _normalscheight - _editpanelheight;
                this.Size = new Size(this.Size.Width, _normalformheight - _editpanelheight);
            }
            return savemode;
        }


        // -----------------------------------------------------------------------
        //  Load Conditions in to the Condition ListView
        // -----------------------------------------------------------------------
        private void LoadConditionList()
        {
            lvSelections.Items.Clear();
            int i;
            SQLCondition cond = new SQLCondition();
            for (i = 0; i <= (SQLFilter.ConditionCount - 1); i++)
            {
                if (SQLFilter.GetCondition(i,ref cond))
                {
                    AddConditionToList(cond);
                }                
            }
        }


        // -----------------------------------------------------------------------
        //  Get the field list from the _fieldsourcequery
        // -----------------------------------------------------------------------
        private int LoadSourceFieldList()
        {
            int i;
            lbFieldName.Items.Clear();

            OleDbConnection connection = new OleDbConnection(_connectionstring);
            if (connection != null)
            {
            connection.Open();                
                string queryString = _fieldsourcequery;
                OleDbCommand command = new OleDbCommand(queryString, connection);
                try
                {
                    OleDbDataReader reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        for (i=0;i<=(reader.FieldCount-1);i++)
                        {
                            string fieldname = reader.GetName(i);
                            lbFieldName.Items.Add(fieldname);
                            _fieldtypes.Add(reader.GetDataTypeName(i).ToString());
                        }
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                }
                connection.Close();
                return 0;
            }
            else
                return -1;
        }


        // -----------------------------------------------------------------------
        //  Build condition set from controls
        // -----------------------------------------------------------------------
        private bool BuildConditionSet(ref SQLCondition cond)
        {
            if ((lbFieldName.SelectedIndex >= 0) && (cbCondition.SelectedIndex >= 0) && (cbValue.Text.Length >= 0))
            {
                cond.fieldtype = _fieldtypes[lbFieldName.SelectedIndex];
                cond.fieldname = lbFieldName.Text;
                cond.condition = cbCondition.Text;
                cond.value = cbValue.Text;
                if (SQLFilter.ConditionCount == 0)
                    cond.logicaloperator = "";
                else
                    if (rbAND.Checked)
                        cond.logicaloperator = "AND";
                    else
                        cond.logicaloperator = "OR";
                return true;
            }
            return false;
        }


        // -----------------------------------------------------------------------
        //  Sets controls according to form mode and state
        // -----------------------------------------------------------------------
        private void SetControls()
        {
            gbOperator.Enabled = ((SQLFilter.ConditionCount > 0) && (cbCondition.SelectedIndex >= 0) && (lbFieldName.SelectedIndex >= 0));
            cbValue.Enabled = ((cbCondition.SelectedIndex >= 0) && (lbFieldName.SelectedIndex >= 0));
            btnAdd.Enabled = ((cbCondition.SelectedIndex >= 0) && (lbFieldName.SelectedIndex >= 0) && (cbValue.Text.Length > 0));
            btnOK.Enabled = (SQLFilter.ConditionCount>0);
        }


        // -----------------------------------------------------------------------
        //  Add condition set to the list
        // -----------------------------------------------------------------------
        private void AddConditionToList(SQLCondition cond)
        {
            ListViewItem lvitem = new ListViewItem();
            lvitem.Text = cond.fieldname; // field name
            lvitem.SubItems.Add(cond.condition); // condition
            lvitem.SubItems.Add(cond.value); // value
            lvitem.SubItems.Add(cond.logicaloperator); // logical operator
            lvSelections.Items.Add(lvitem);
        }


        // -----------------------------------------------------------------------
        //  Add a field to the list
        // -----------------------------------------------------------------------
        private void LoadValuesItems()
        {
            if (lbFieldName.SelectedIndex < 0) return;
            cbValue.Items.Clear();
            cbValue.Text = "";
            List<string> ValueList = new List<string>();
            try
            {
                OleDbConnection connection = new OleDbConnection(_connectionstring);
                if (connection != null)
                {
                    connection.Open();
                    string tablename = _layers.CWTableName(_layerindex);
                    if (tablename.Length == 0) return;
                    string queryString = "SELECT DISTINCT [" + lbFieldName.Text + "] FROM [" + tablename + "]";
                    OleDbCommand command = new OleDbCommand(queryString, connection);
                    try
                    {
                        OleDbDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            cbValue.Items.Add(reader[0].ToString());
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                    }
                    connection.Close();
                    return;
                }
                else
                    return;
            }
            finally
            {
            }
        }            


        // -----------------------------------------------------------------------
        //  OK Button Click
        // -----------------------------------------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {
            // Generate the Filter Class and Leave
            if (_editmode) // if editing an existing filter, apply the changes made
            {
                if (_originalfilter != null)
                    _originalfilter.Assign(SQLFilter);
            }
        }

        // -----------------------------------------------------------------------
        //  Add Button Click
        // -----------------------------------------------------------------------
        private void btnAdd_Click(object sender, EventArgs e)
        {
            SQLCondition cond = new SQLCondition();
            if (BuildConditionSet(ref cond))
            {
                if (SQLFilter.AddCondition(cond))
                {
                    AddConditionToList(cond);
                    cbCondition.SelectedIndex = 0;
                    cbValue.Text = "";                    
                }
            }
            SetControls();
        }


        // -----------------------------------------------------------------------
        //  Field Selection Changed
        // -----------------------------------------------------------------------
        private void lbFieldName_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Load the distinct values of the selected field
            LoadValuesItems();
            SetControls();
        }


        // -----------------------------------------------------------------------
        //  OK Button Click
        // -----------------------------------------------------------------------
        private void btnOK_Click(object sender, EventArgs e)
        {
            // Generate the Filter Class and Leave
            if (_editmode) // if editing an existing filter, apply the changes made
            {
                SQLFilter.FilterName = tbFilterName.Text;
                SQLFilter.FilterDescription = tbFilterDescription.Text;
                if (_originalfilter != null)
                    _originalfilter.Assign(SQLFilter);
            }
        }

        private void cbValue_TextChanged(object sender, EventArgs e)
        {
            SetControls();
        }

    }
}
