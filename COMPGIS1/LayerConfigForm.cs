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
    public partial class LayerConfigForm : Form
    {
        private bool _initok = false;
        private ArcEngineClass _GISdata;
        private List<string> _layerlist = new List<string>();

        //---------------------------------------------------
        // Constructor
        //---------------------------------------------------
        public LayerConfigForm(ArcEngineClass gisData)
        {
            InitializeComponent();
            _GISdata = gisData;
            _initok = (gisData != null);
        }


        //-----------------------------------------------------------------------------------
        // Form Load Event
        //-----------------------------------------------------------------------------------
        private void LayerConfigForm_Load(object sender, EventArgs e)
        {
            if (!_initok)
            {
                MessageBox.Show("Error loading the Layer Configuration Form!");
                this.Close();
            }
            LoadKeyFieldSettings();
            BuildLayerMapGrid();
        }

        //---------------------------------------------------
        // BuildLayerMapGrid
        //---------------------------------------------------
        private void BuildLayerMapGrid()
        {
            // Get the layers to put in the dropdown
            if (_GISdata.layer_getLayerList(ref _layerlist))
            {
                // Load the Layers in to the Mapped Layer Dropdown in the grid
                ColMappedLayer.Items.Clear();
                int i;
                for (i = 0; i <= (_layerlist.Count - 1); i++)
                {
                    ColMappedLayer.Items.Add(_layerlist[i]);
                }

                // Load the LayerMap grid with the layer records
                gridLayerMap.Rows.Clear();
                for (i = 0; i <= (_GISdata._layers.Count - 1); i++)
                {
                    int rowId = gridLayerMap.Rows.Add();
                    DataGridViewRow row = gridLayerMap.Rows[rowId];
                    // Load Grid
                    LayerDef ldef = new LayerDef();
                    _GISdata._layers.GetLayerDefOrdinal(i, ref ldef);
                    row.Cells[0].Value = _GISdata._layers.GetLayerTypeName(ldef.SystemUse);
                    row.Cells[1].Value = "";
                    if (ColMappedLayer.Items.IndexOf(ldef.layername) >= 0)
                    {
                        row.Cells[1].Value = ldef.layername;
                    }
                }
            }
        }


        //-----------------------------------------------------------------------------------
        // Load the Attribute table Key Fields xref
        //-----------------------------------------------------------------------------------
        private void LoadKeyFieldSettings()
        {
            gridKeyFields.Rows.Clear();
            List<string> FieldList = new List<string>();
            int i;
            LayerDef ldef = new LayerDef();
            for (i = 0; i <= (_GISdata._layers.Count - 1); i++)
            {
                int rowId = gridKeyFields.Rows.Add();
                DataGridViewRow row = gridKeyFields.Rows[rowId];
                // Load Grid
                _GISdata._layers.GetLayerDefOrdinal(i, ref ldef);
                row.Cells[0].Value = _GISdata._layers.GetLayerTypeName(ldef.SystemUse);
                //row.Cells[0].Value = ldef.layername;
                row.Cells[1].Value = ldef.GIS_Attribute_Index_Field;
            }
        }


        //.................................................
        //.................................................
        //.................................................
        private bool SelectCWIndexField(int index)
        {
            if (index < 0)
                return false;
            List<string> FieldList = new List<string>();
            // Get Layer Def
            LayerDef ldef = new LayerDef();
            _GISdata._layers.GetLayerDefOrdinal(index, ref ldef);
            // Get selected item to pass to selection form
            int defvalue = -1;
            if (_GISdata.GetLayerAttributeFields(ldef.layername, ref FieldList) > 0)
            {
                string curvalue = gridKeyFields.Rows[index].Cells[1].Value.ToString();
                if (curvalue.Length > 0)
                {
                    defvalue = FieldList.IndexOf(curvalue);
                }
                SelectFromStringList selform = new SelectFromStringList(ref FieldList, defvalue, string.Format("Select Key [{0}]", ldef.layername));
                if (selform.Initok)
                {
                    if ((selform.ShowDialog() == DialogResult.OK) && (selform.Selecteditem >= 0) && (selform.Selecteditem != defvalue))
                    {
                        gridKeyFields.Rows[index].Cells[1].Value = FieldList[selform.Selecteditem];
                        gridKeyFields.Update();
                        return true;
                    }
                }
            }
            return false;
        }


        //-----------------------------------------------------------------------------------
        // User clicked OK button
        //-----------------------------------------------------------------------------------
        private bool SaveSettings()
        {
            // Save Selected Layer Mappings
            int i;
            try
            {
                LayerDef ldef = new LayerDef();
                for (i = 0; i <= (_GISdata._layers.Count - 1); i++)
                {
                    _GISdata._layers.GetLayerDefOrdinal(i, ref ldef);
                    string newvalue;

                    // Save Layer Name if changed
                    newvalue = gridLayerMap.Rows[i].Cells[1].Value.ToString();
                    if (string.Compare(ldef.layername, newvalue) != 0)
                    {
                        _GISdata._cwdata.SetLayerName((int)ldef.SystemUse, newvalue);
                        ldef.layername = newvalue;
                    }
                    // Save the Key Fields if changed
                    newvalue = gridKeyFields.Rows[i].Cells[1].Value.ToString();
                    if (string.Compare(ldef.GIS_Attribute_Index_Field, newvalue) != 0)
                    {
                        _GISdata._cwdata.SetLayerCWKeyField((int)ldef.SystemUse, newvalue);
                        ldef.GIS_Attribute_Index_Field = newvalue;
                    }
                }               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to save the Settings!");
                return false;
            }
            return true;
        }



        //-----------------------------------------------------------------------------------
        // User double-clicked on the Key Field grid
        //-----------------------------------------------------------------------------------
        private void gridKeyFields_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            SelectCWIndexField(e.RowIndex);
        }

        //-----------------------------------------------------------------------------------
        // User clicked on the Select button
        //-----------------------------------------------------------------------------------
        private void btnSelect_Click(object sender, EventArgs e)
        {
            SelectCWIndexField(gridKeyFields.CurrentCell.RowIndex);
        }

        //-----------------------------------------------------------------------------------
        // User clicked OK button
        //-----------------------------------------------------------------------------------
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (SaveSettings())
            {
                this.Close();
            }
        }






    }
}
