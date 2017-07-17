using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Timers;
using System.Windows.Forms;

namespace COMPGIS1
{


    public partial class BPDBUpdates : Form
    {

        private ArcEngineClass _GISdata;
        private ArcEngineBP _BPGISdata;
        private bool _updatesrequired = false;
        private bool _updateriskranking = false;
        private bool _userquit = false;
        private bool _promptuser = false;

        //---------------------------------------------------
        // Constructor
        //---------------------------------------------------
        public BPDBUpdates(ArcEngineBP BPGISdata)
        {
            InitializeComponent();
            _BPGISdata = BPGISdata;
            _GISdata = (ArcEngineClass) BPGISdata;
        }

        public BPDBUpdates(ArcEngineClass GISdata,bool promptUser)
        {
            InitializeComponent();
            _GISdata = GISdata;
            _promptuser = promptUser;
        }

        //---------------------------------------------------
        // For Load Event
        //---------------------------------------------------
        private void DBUpdates_Load(object sender, EventArgs e)
        {
            if (_BPGISdata != null)
            {
                _updatesrequired = (_BPGISdata.CheckForFieldUpdates());
                if (_promptuser)
                {
                    MessageBox.Show("There are no updates required!");
                }
            }
        }


        //---------------------------------------------------
        // Do Updates
        //---------------------------------------------------
        public void PerformUpdates()
        {
            int idx = 0;
            if (!_updatesrequired) return;
            if (_GISdata != null)
            {
                // Risk Ranking Updates
                if (_GISdata._cwdata.linesectionCheckForUpdates("R"))
                {
                    clbStatus.Items.Add("Line Section Risk Ranking");

                    // Do the RR update
                    linesectionUpdateRiskRanking();

                    // Done with RR update
                    clbStatus.SetItemChecked(idx, true);
                    idx++;

                }

                // Add any others

                Close();
            }
        }


        //-----------------------------------------------------------------------------------
        // Get Field Data from the Reader
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
            catch (Exception ex)
            {
            }
            finally
            {
            }
            return retval;
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
            catch (Exception ex)
            {
            }
            finally
            {
            }
            return retval;
        }



        //------------------------------------------------------------------------------
        // Get Items Needing Update from CompWorks
        // Returns a string with the characters representing the types of updates found
        //------------------------------------------------------------------------------
        public bool linesectionUpdateRiskRanking()
        {
            bool retval = false;

            // Check the "Risk Ranking" [R]
            OleDbConnection connection = new OleDbConnection(_GISdata._cwdata._connectionstring);
            OleDbDataReader reader;
            if (connection != null)
            {
                try
                {
                    connection.Open();
                    string cmdStr = "SELECT [Dbf Line Section].[Zone ID], [Dbf Line Section].[Line Section], [Dbf Line Section].SusRanking, [Dbf Line Section].ConRanking, [Dbf Line Section].[Overall Ranking], [Dbf Line Section].UpdateFlags FROM [Dbf Line Section] WHERE ((([Dbf Line Section].UpdateFlags) Like '%R%'))";
                    OleDbCommand command = new OleDbCommand(cmdStr, connection);
                    if (command != null)
                    {
                        double progress = 0;
                        int rowCurrent = 0;
                        int rowCount = (int)command.ExecuteScalar();
                        double rowCountdbl = Convert.ToDouble(rowCount);
                        if (rowCount == 0) return retval;
                        reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            rowCurrent++;
                            int ZoneID = GetFieldDataInt("Zone ID", ref reader);
                            int Susc = GetFieldDataInt("SusRanking", ref reader);
                            int Cons = GetFieldDataInt("ConRanking", ref reader);
                            int Overall = GetFieldDataInt("Overall Ranking", ref reader);

                            if ((ZoneID > 0) && (Susc >= 0) && (Cons >= 0) && (Overall >= 0))
                            {
                                _BPGISdata.linesectionUpdateRR(ZoneID, Overall, Susc, Cons);
                                _GISdata._cwdata.linesectionRemoveUpdateFlag("R", ZoneID);
                            }

                            progress = ((rowCurrent/rowCountdbl)*100);
                            pbMain.Value = Convert.ToInt32(progress);
                            rowCurrent++;
                            Application.DoEvents();
                            Refresh();
                            pbMain.Refresh();

                        if (_userquit)
                        {
                            reader.Close();
                            connection.Close();
                            return retval;
                        }

                        }
                        reader.Close();
                    }                   
                }
                catch (Exception ex)
                {                    
                }
                connection.Close();
            }
            return retval;
        }

        //------------------------------------------------------------------------------
        // User clicked Cancel
        //------------------------------------------------------------------------------
        private void btnClose_Click(object sender, EventArgs e)
        {
            _userquit = true;
        }




    
    }

}
