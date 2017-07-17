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

    public partial class SelectFromStringList : Form
    {

        private int _defaultitem;
        private bool _initok = false;
        private int _selecteditem = -1;
        private List<string> _stringlist = null;

        public bool Initok
        {
            get { return _initok; }
        }

        public int Selecteditem
        {
            get { return _selecteditem; }
        }


        //..........................................................................................
        // Constructor
        //..........................................................................................
        public SelectFromStringList(ref List<string> stringList, int defaultItem, string formCaption)
        {
            InitializeComponent();
            if (formCaption.Length>0)
                this.Text = formCaption;
            _defaultitem = defaultItem;
            lbItems.Items.Clear();
            if (stringList != null)
            {
                _stringlist = stringList;
                _initok = true;
            }
        }

        //..........................................................................................
        // Form Load Event
        //..........................................................................................
        private void SelectFromStringList_Load(object sender, EventArgs e)
        {
            int i;
            if (_initok)
            {
                for (i = 0; i <= (_stringlist.Count - 1); i++)
                {
                    lbItems.Items.Add(_stringlist[i]);
                }
                _initok = true;
                if (lbItems.Items.Count > 0)
                    if (_defaultitem >= 0)
                    {
                        lbItems.SelectedIndex = _defaultitem;
                    }
                    else
                        lbItems.SelectedIndex = 0;
            }
        }


        //..........................................................................................
        // Set Controls on form
        //..........................................................................................
        public void SetControls()
        {
            btnOK.Enabled = (_initok && (lbItems.SelectedIndex>=0));

        }

        //..........................................................................................
        // Item Selected
        //..........................................................................................
        private void lbItems_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetControls();
        }

        //..........................................................................................
        // Item Selected
        //..........................................................................................
        private void SelectCurrentItem()
        {
            if (_initok && (lbItems.SelectedIndex >= 0))
            {
                _selecteditem = lbItems.SelectedIndex;
                this.Close();
            }
        }

        //..........................................................................................
        // User Selected Item and Clicked OK
        //..........................................................................................
        private void btnOK_Click(object sender, EventArgs e)
        {
            SelectCurrentItem();
        }

        //..........................................................................................
        // User Double clicked on Item
        //..........................................................................................
        private void lbItems_DoubleClick(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            SelectCurrentItem();
        }



    }
}
