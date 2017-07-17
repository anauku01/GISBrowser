using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace COMPGIS1
{
    public partial class BPTestStationDetails : Form
    {

        private string _connectionstring = "";
        private int _teststationid = 0;

        //-----------------------------------------------------------------------------------
        // Contructor
        //-----------------------------------------------------------------------------------
        public BPTestStationDetails(int TestStationID, string connectionstring)
        {
            InitializeComponent();
            _connectionstring = connectionstring;
            
            if (_connectionstring.Length > 0)
            {
                _teststationid = TestStationID;
                LoadFormData(_teststationid);
                this.Text = "Test Station";
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
        public int LoadFormData(int TestStationID)
        {
            int rectifierid = -1;
            string queryString = "";
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);

            if (connection != null)
            {
                connection.Open();

                queryString = string.Format("SELECT [Dbf Stations].TestStationID, [Dbf Stations].[Test Station], [Dbf Stations].[In Service Date], [Dbf Stations].Location, [Dbf Stations].Remarks, [Dbf Stations].CompanyID, [Dbf Stations].GUID, [Dbf Setup -  Data].ValueData AS TestStatType FROM [Dbf Stations] INNER JOIN [Dbf Setup -  Data] ON [Dbf Stations].[Test Station Type] = [Dbf Setup -  Data].SetupID WHERE ((([Dbf Stations].TestStationID)={0}))", TestStationID);
                command = new OleDbCommand(queryString, connection);
                try
                {
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        // Fill in controls from the Lines Table
                        tbTestStation.Text = GetFieldDataStr("Test Station", ref reader);
                        tbInServiceDate.Text = GetFieldDataStr("In Service Date", ref reader);
                        tbTestStationType.Text = GetFieldDataStr("TestStatType", ref reader);
                        tbRemarks.Text = GetFieldDataStr("Remarks", ref reader);
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
        // DONE
        return 0;
        } // loadformdata 


















    }

}
