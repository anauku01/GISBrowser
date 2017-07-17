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
    public partial class LoginAdmin : Form
    {

        private bool _loginok = false;
        private int _pwlen = 7;
        public bool LoginOK { get { return _loginok; } set { } }
        
        #region class constructor
        public LoginAdmin()
        {
            InitializeComponent();
            this.CancelButton = btnCancel;
            this.AcceptButton = btnOK;
        }
        #endregion

        //---------------------------------------------------
        // Form Load Method
        //---------------------------------------------------
        private void LoginAdmin_Load(object sender, EventArgs e)
        {
            // Set the focus to the Edit Control
            edPW.Focus();
        }

        //---------------------------------------------------
        // Set OK Button enabled based upon data in edit
        //---------------------------------------------------
        private void SetControls()
        {
            btnOK.Enabled = (edPW.Text.Length > 0);
        }

        //---------------------------------------------------
        // Form Closing Event 
        // :: Do not close unless password has been entered
        //---------------------------------------------------
        private void LoginAdmin_FormClosing(object sender, FormClosingEventArgs e)
        {
            const string errmsg = "Password Missing or Incorrect!";
            if (this.DialogResult == DialogResult.OK)
            {
                string pw = edPW.Text;
                bool isok = true;
                if (pw.Length == _pwlen)
                {
                    int i;
                    for (i = 0; i <= (pw.Length - 1); i++)
                    {
                        switch (i)
                        {
                            case 0:
                                isok = (pw.Substring(0, 1) == "w");
                                break;
                            case 1:
                                isok = (pw.Substring(1, 1) == "a");
                                break;
                            case 2:
                                isok = (pw.Substring(2, 1) == "i");
                                break;
                            case 3:
                                isok = (pw.Substring(3, 1) == "o");
                                break;
                            case 4:
                                isok = (pw.Substring(4, 1) == "u");
                                break;
                            case 5:
                                isok = (pw.Substring(5, 1) == "r");
                                break;
                            case 6:
                                isok = (pw.Substring(6, 1) == "u");
                                break;
                        }
                    }
                    e.Cancel = (!isok);
                    _loginok = isok;
                    if (!isok)
                        lblError.Text = errmsg;
                }
                else
                {
                    e.Cancel = true;
                    lblError.Text = errmsg;
                }                
            }
        }

        //---------------------------------------------------
        // PW Edit Change Event
        //---------------------------------------------------
        private void edPW_TextChanged(object sender, EventArgs e)
        {
            SetControls();
            lblError.Text = "";
        }


    }
}
