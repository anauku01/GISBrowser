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
    public partial class findUserSelectBP : Form
    {

        private string _connectionstring = "";
        private int _rectifierid = 0;
        private LayerDefs _ldefs;

        //-----------------------------------------------------------------
        // Constructor
        //-----------------------------------------------------------------
        public findUserSelectBP(LayerDefs ldefs, string connectionstring)
        {
            InitializeComponent();
            _ldefs = ldefs;
        }

        //-----------------------------------------------------------------
        // Form Load Event
        //-----------------------------------------------------------------
        private void findUserSelectBP_Load(object sender, EventArgs e)
        {

        }




    }
}
