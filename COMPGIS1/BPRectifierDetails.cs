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
    public partial class BPRectifierDetails : Form
    {

        private string _connectionstring = "";
        private int _rectifierid = 0;


        //-----------------------------------------------------------------------------------
        // Contructor
        //-----------------------------------------------------------------------------------
        public BPRectifierDetails(int RectifierID, string connectionstring)
        {
            InitializeComponent();
            _connectionstring = connectionstring;
            if (_connectionstring.Length > 0)
            {
                _rectifierid = RectifierID;
                LoadFormData(_rectifierid);
                this.Text = "Rectifier Details";
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
        public int LoadFormData(int RectifierID)
        {
            int rectifierid = -1;
            string queryString = "";
            OleDbCommand command;
            OleDbDataReader reader;
            OleDbConnection connection = new OleDbConnection(_connectionstring);

            if (connection != null)
            {
                connection.Open();
                queryString = string.Format("SELECT [Dbf Rectifiers].RectifierID, [Dbf Rectifiers].Rectifier, [Dbf Rectifiers].[In Service Date], [Dbf Rectifiers].Manufacturer, [Dbf Rectifiers].Model, [Dbf Rectifiers].[Operating Amps Out], " +
                                            "[Dbf Rectifiers].[Operating Voltage Out], [Dbf Rectifiers].[Power Source], [Dbf Rectifiers].[Rated Amps Out], [Dbf Rectifiers].[Rated Voltage Out], [Dbf Rectifiers].[Stack Type], [Dbf Rectifiers].[Replaced By Date], " +
                                            "[Dbf Rectifiers].[Negative Counts], [Dbf Rectifiers].[Anodes Counts], [Dbf Rectifiers].Remarks, [Dbf Rectifiers].CompanyID, [Dbf Rectifiers].GUID, [Dbf Rectifiers].PrevRectSurvResOp, [Dbf Rectifiers].Operational, " +
                                            "[Dbf Rectifiers].DateOfLastInspection, [Dbf Rectifiers].MapFeatureID FROM [Dbf Rectifiers] WHERE ((([Dbf Rectifiers].RectifierID)={0}));", RectifierID);
                command = new OleDbCommand(queryString, connection);
                try
                {
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        // Fill in controls from the Lines Table
                        tbRectifier.Text = GetFieldDataStr("Rectifier", ref reader);
                        tbInServiceDate.Text = GetFieldDataStr("In Service Date", ref reader);
                        tbManufacturer.Text = GetFieldDataStr("Manufacturer", ref reader);
                        tbModel.Text = GetFieldDataStr("Model", ref reader);
                        tbOperAmps.Text = GetFieldDataStr("Operating Amps Out", ref reader);
                        tbOperVolts.Text = GetFieldDataStr("Operating Voltage Out", ref reader);
                        tbPowerSource.Text = GetFieldDataStr("Power Source", ref reader);
                        tbRatedAmps.Text = GetFieldDataStr("Rated Amps Out", ref reader);
                        tbRatedVolts.Text = GetFieldDataStr("Rated Voltage Out", ref reader);
                        tbStackType.Text = GetFieldDataStr("Stack Type", ref reader);
                        tbReplaceByDate.Text = GetFieldDataStr("Replaced By Date", ref reader);
                        tbLastInsp.Text = GetFieldDataStr("DateOfLastInspection", ref reader);
                        tbAnodeCount.Text = GetFieldDataStr("Anodes Counts", ref reader);
                        tbComments.Text = GetFieldDataStr("Remarks", ref reader);
                        // Boolean Fields
                        cbOperational.Checked = (Convert.ToBoolean(GetFieldDataStr("Operational", ref reader)));
                        rectifierid = GetFieldDataInt("RectifierID", ref reader);

                    }
                reader.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                    return -1;
                }

                connection.Close();

                //...............
                // Grid Rectifier Readings
                //...............
                //if (rectifierid > 0)
                //{
                //    queryString =
                //        "SELECT [Dbf Rectifiers].RectifierID, [Dbf Rectifiers].Rectifier, [Dbf Rectifiers Readings].[Reading Date], [Dbf Rectifiers Readings].Voltage, " +
                //        "[Dbf Rectifiers Readings].Current FROM [Dbf Rectifiers] INNER JOIN [Dbf Rectifiers Readings] ON [Dbf Rectifiers].RectifierID = [Dbf Rectifiers Readings].RectifierID " +
                //        "WHERE ((([Dbf Rectifiers].RectifierID)={0})) ORDER BY [Dbf Rectifiers Readings].[Reading Date] DESC";

                //    queryString = string.Format(queryString, rectifierid);
                //    OleDbDataAdapter dAdapter = new OleDbDataAdapter(queryString, _connectionstring);
                //    //OleDbCommandBuilder cBuilder = new OleDbCommandBuilder(dAdapter);
                //    DataTable dTable = new DataTable();
                //    //fill the DataTable
                //    dAdapter.Fill(dTable);
                //    //the DataGridView
                //    //DataGridView dgView = new DataGridView();
                //    BindingSource bSource = new BindingSource();
                //    bSource.DataSource = dTable;
                //    gridReadings.DataSource = bSource;
                //    gridReadings.Columns["RectifierID"].Visible = false;
                //}

            connection.Close();
            } // if connection!=null
        // DONE
        return 0;
        } // loadformdata 





    }
}
