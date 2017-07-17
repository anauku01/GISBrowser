using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing;
using System.Data.OleDb;

namespace COMPGIS1
{
    class RadarPlotGenerator
    {

        private string _connectionstring = "";
        private Chart _chart;
        private string _destbitmapfile;
        private string _titletop1;
        private string _titletop2;
        private string _titlebottom;
        private double _Ca;
        private double _Mg2;
        private double _K;
        private double _Na;
        private double _Cl;
        private double _S042;
        private double _CaExt;
        private double _Mg2Ext;
        private double _KExt;
        private double _NaExt;
        private double _ClExt;
        private double _S042Ext;
        private int _soil_chemical_results_id;
        private bool _plotcomplete = false;
        public bool plotcomplete { set{} get { return _plotcomplete; } }

        //--------------------------------------------------------------------------------------------------------------------------------------------
        // Constructor - Overload Option 1
        //--------------------------------------------------------------------------------------------------------------------------------------------
        public RadarPlotGenerator(Chart chart, string titletop1, string titletop2, string titlebottom, double Ca, double Mg2, double K, double Na, double Cl, double S042)
        {
            if (chart == null) return;
            _chart = chart;
            _titletop1 = titletop1;
            _titletop2 = titletop2;
            _titlebottom = titlebottom;
            _Ca = Ca;
            _Mg2 = Mg2;
            _K = K;
            _Na = Na;
            _Cl = Cl;
            _S042 = S042;
            _plotcomplete = (LoadChart() == 0);
        }

        //--------------------------------------------------------------------------------------------------------------------------------------------
        // Constructor - Overload Option 2
        //--------------------------------------------------------------------------------------------------------------------------------------------
        public RadarPlotGenerator(Chart chart, string titletop1, string titletop2, string titlebottom, int Soil_Chemical_Results_ID, string ConnectionString)
        {
            if (chart == null) return;
            _chart = chart;
            _titletop1 = titletop1;
            _titletop2 = titletop2;
            _titlebottom = titlebottom;
            _soil_chemical_results_id = Soil_Chemical_Results_ID;
            _connectionstring = ConnectionString;
            if (Calculate_Values(Soil_Chemical_Results_ID) == true)
            {
                _plotcomplete = (LoadChart() == 0);
            }            
        }

        //--------------------------------------------------------------------------------------------------------------------------------------------
        //--------------------------------------------------------------------------------------------------------------------------------------------
        private bool GetMolarEquivalentValues(string Element, ref double MolarMass, ref double Charge)
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
                    queryString =
                        "SELECT [Dbf Setup - Molar Equivalents].Element_Name, [Dbf Setup - Molar Equivalents].Molar_Mass, [Dbf Setup - Molar Equivalents].Charge FROM [Dbf Setup - Molar Equivalents] " +
                        "WHERE ([Dbf Setup - Molar Equivalents].Element_Name = \"" + Element + "\")";
                    command = new OleDbCommand(queryString, connection);
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        if (reader.FieldCount == 3)
                        {
                            MolarMass = Convert.ToDouble(reader[1].ToString());
                            Charge = Convert.ToDouble(reader[2].ToString());
                        }
                        else
                            return false;
                    }
                    else
                        return false;
                }
                finally
                {
                }
                return true;
            }
            else
                return false;
        }


        //--------------------------------------------------------------------------------------------------------------------------------------------
        // Returns the Extractable Values from the chemical analysis results
        //--------------------------------------------------------------------------------------------------------------------------------------------
        private bool GetExtractableValues()
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

                    queryString =   "SELECT Soil_Chemical_Results.Soil_Chemical_Results_ID, Soil_Chemical_Results.Extractable_Calcium_mg_kg, Soil_Chemical_Results.Extractable_Magnesium_mg_kg, Soil_Chemical_Results.Extractable_Potassium_mg_kg, Soil_Chemical_Results.Extractable_Sodium_mg_kg, Soil_Chemical_Results.Extractable_Chloride_mg_kg, Soil_Chemical_Results.Extractable_Sulfate_mg_kg FROM Soil_Chemical_Results " + 
                                    "WHERE (((Soil_Chemical_Results.Soil_Chemical_Results_ID)=" + _soil_chemical_results_id.ToString() + "));";
                    command = new OleDbCommand(queryString, connection);
                    reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        if (reader.FieldCount == 7)
                        {
                            _CaExt = Convert.ToDouble(reader[1].ToString());
                            _Mg2Ext = Convert.ToDouble(reader[2].ToString());
                            _KExt = Convert.ToDouble(reader[3].ToString());
                            _NaExt = Convert.ToDouble(reader[4].ToString());
                            _ClExt = Convert.ToDouble(reader[5].ToString());
                            _S042Ext = Convert.ToDouble(reader[6].ToString());
                        }
                        else
                            return false;
                    }
                    else
                        return false;
                }
                finally
                {
                }
                return true;
            }
            else
                return false;
        }


        //--------------------------------------------------------------------------------------------------------------------------------------------
        // Calculates the ION Concentration values in the polar plot
        //--------------------------------------------------------------------------------------------------------------------------------------------
        private bool Calculate_Values(int Soil_Chemical_Results_ID)
        {
            double MolarMass = 0;
            double Charge = 0;
            try
            {
                // Get the extractable values from the analysis results record
                if (GetExtractableValues() != true) return false;
                // Calculate the plot values (ION Concentrations)
                if (GetMolarEquivalentValues("Ca2", ref MolarMass, ref Charge) && (Charge > 0) && (MolarMass > 0))
                    _Ca = (_CaExt * Charge)/MolarMass;
                if (GetMolarEquivalentValues("Mg2", ref MolarMass, ref Charge) && (Charge > 0) && (MolarMass > 0))
                    _Mg2 = (_Mg2Ext * Charge)/MolarMass;
                if (GetMolarEquivalentValues("K", ref MolarMass, ref Charge) && (Charge > 0) && (MolarMass > 0))
                    _K = (_KExt * Charge)/MolarMass;
                if (GetMolarEquivalentValues("Na", ref MolarMass, ref Charge) && (Charge > 0) && (MolarMass > 0))
                    _Na = (_NaExt * Charge)/MolarMass;
                if (GetMolarEquivalentValues("Cl", ref MolarMass, ref Charge) && (Charge > 0) && (MolarMass > 0))
                    _Cl = (_ClExt * Charge)/MolarMass;
                if (GetMolarEquivalentValues("SO4", ref MolarMass, ref Charge) && (Charge > 0) && (MolarMass > 0))
                    _S042 = (_S042Ext * Charge)/MolarMass;
            }
            catch (Exception ex)
            {
                return false;
            }
            
            return true;
        }


        //--------------------------------------------------------------------------------------------------------------------------------------------
        //--------------------------------------------------------------------------------------------------------------------------------------------
        private void add_chart_series(ref System.Windows.Forms.DataVisualization.Charting.Chart chart, ref int dpoint, double datavalue, string axislabel)
        {
            chart.Series[0].Points.Add(datavalue);
            chart.Series[0].BorderWidth = 2;
            chart.Series[0].Color = Color.Navy;
            chart.Series[0].Points[dpoint].MarkerSize = 8;
            chart.Series[0].Points[dpoint].MarkerStyle = MarkerStyle.Square;
            chart.Series[0].Points[dpoint].MarkerColor = Color.Navy;
            chart.Series[0].Points[dpoint].AxisLabel = axislabel;
            dpoint++;
        }

        //--------------------------------------------------------------------------------------------------------------------------------------------
        //--------------------------------------------------------------------------------------------------------------------------------------------
        private int LoadChart()
        {
            int dpoint = 0;
            try
            {
                _chart.Legends.Clear();
                _chart.Series[0].ChartType = SeriesChartType.Radar;
                _chart.Series[0]["AreaDrawingStyle"] = "Polygon";
                _chart.Series[0]["RadarDrawingStyle"] = "Line";
                _chart.Series[0].Points.Clear();
                _chart.ChartAreas[0].Axes[0].LabelStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 10,
                                                                                       System.Drawing.FontStyle.Bold);
                //            _chart.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8, System.Drawing.FontStyle.Bold);
                _chart.ChartAreas[0].AxisY.LabelStyle.Format = "E";
                _chart.Series[0].Points.Clear();
                add_chart_series(ref _chart, ref dpoint, _Ca, "Ca2+");
                add_chart_series(ref _chart, ref dpoint, _Mg2, "Mg2+");
                add_chart_series(ref _chart, ref dpoint, _K, "K+");
                add_chart_series(ref _chart, ref dpoint, _Na, "Na+");
                add_chart_series(ref _chart, ref dpoint, _Cl, "Cl-");
                add_chart_series(ref _chart, ref dpoint, _S042, "S042-");
                _chart.Titles.Clear();
                // Add titles if existing
                _chart.Titles.Add(new Title("Chart Title", Docking.Top, new Font("Tahoma", 12, FontStyle.Bold),
                                            Color.Black));
                _chart.Titles[0].Visible = (_titletop1.Length > 0);
                _chart.Titles[0].Text = _titletop1;
                _chart.Titles.Add(new Title("Chart Sub-Title", Docking.Top, new Font("Tahoma", 10, FontStyle.Regular),
                                            Color.Black));
                _chart.Titles[1].Visible = (_titletop2.Length > 0);
                _chart.Titles[1].Text = _titletop2;
                _chart.Titles.Add(new Title("Bottom", Docking.Bottom, new Font("Tahoma", 12, FontStyle.Regular),
                                            Color.Black));
                _chart.Titles[2].Visible = (_titlebottom.Length > 0);
                _chart.Titles[2].Text = _titlebottom;
            }
            catch
            {
                return 1;
            }
            return 0;
        }

        //--------------------------------------------------------------------------------------------------------------------------------------------
        //--------------------------------------------------------------------------------------------------------------------------------------------
        public void SaveChartToJPG(string destbitmapfile)
        {
            if (destbitmapfile.Length == 0) return;
            _destbitmapfile = destbitmapfile;
            _chart.SaveImage(_destbitmapfile, System.Drawing.Imaging.ImageFormat.Jpeg);
        }


    }
}
