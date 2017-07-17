using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.esriSystem;


namespace COMPGIS1
{
    public partial class BPLineSectionLayerSettings : Form
    {
        private string _styleFilePath;
        private bool _styleFileFound = false;
        private LayerDef _layerDef;
        private bool _wasmodified = false;

        // Statistics
        private int _totalLength = 0;
        private List<string> _unitlengths = new List<string>();

        // Display Values
        private double _minriskvalue;
        private double _maxriskvalue;
        private bool _displayModified = false;
        private int _displaymode = 0;
        private int _riskfieldtouse = 0;
        private SimpleRenderer m_SimpleRenderer;
        private ClassBreaksRenderer m_classBreaksRenderer;
        private IStyleGalleryItem m_styleGalleryItem;
        private IFeatureLayer m_featureLayer;
        private ArcEngineClass _GISdata;
        private string _selectedcolorrampname = "";


        public bool DisplayModified
        {
            get { return _displayModified; }
        }

        public ClassBreaksRenderer LineSectionRender
        {
            get { return m_classBreaksRenderer; }
        }

        public string SelectedColorRampName
        {
            get { return _selectedcolorrampname; }
        }

        public SimpleRenderer SimpleRenderer
        {
            get { return m_SimpleRenderer; }
        }

        public int Displaymode
        {
            get { return _displaymode; }
        }

        //.................................................
        //.................................................
        //.................................................
        public BPLineSectionLayerSettings(LayerDef layerDef, ArcEngineClass GISdata)
        {
            InitializeComponent();
            cbRiskRankingChoices.SelectedIndex = 0;
            if (layerDef != null)
            {
                _layerDef = layerDef;
            }
            if (GISdata != null)
            {
                _GISdata = GISdata;
            }

        }

        //.................................................
        //.................................................
        //.................................................
        private void LineSectionLayerSettings_Load(object sender, EventArgs e)
        {
            if (_layerDef != null)
            {
                lblLayerName.Text = _layerDef.layername;
                lblLayerDesc.Text = _layerDef.Description;
            }

            // Load the Color Ramps

            _styleFilePath = Path.ChangeExtension(Application.ExecutablePath, ".serverStyle");

            if (File.Exists(_styleFilePath))
            {
                //Load the ESRI.ServerStyle file into the SymbologyControl
                axSymbologyControl1.LoadStyleFile(_styleFilePath);
                //Set the style class
                axSymbologyControl1.StyleClass = ESRI.ArcGIS.Controls.esriSymbologyStyleClass.esriStyleClassColorRamps;
                //Select the color ramp item
                axSymbologyControl1.GetStyleClass(axSymbologyControl1.StyleClass).SelectItem(0);
                _styleFileFound = true;
            }
            else
            {
                if (rbColorRamp.Checked)
                    rbByLayer.Checked = true;
                rbColorRamp.Enabled = false;
                _styleFileFound = false;
                lblColorRampName.Text = "Color Ramp Style File not Found!";
            }


            //Get the Layer Info needed
            IGeoFeatureLayer geofeaturelayer = null;
            IMap map = _GISdata._activeview.FocusMap;
            for (int i = 0; i <= map.LayerCount - 1; i++)
            {
                if (map.get_Layer(i).Name == _layerDef.layername)
                {
                    geofeaturelayer = (IGeoFeatureLayer)map.get_Layer(i);
                    break;                    
                }
            }
            if (geofeaturelayer == null) return;
            m_featureLayer = geofeaturelayer;

            // Initialize Default Display Renderer
            m_SimpleRenderer = new SimpleRendererClass();
            AssignDefaultRenderer();

            // See what current settings are and set them on the form
            _displaymode = _GISdata._cwdata.Configuration.BPLSDisplayMode;
            _selectedcolorrampname = _GISdata._cwdata.Configuration.BPSavedColorRampName;
            if (_styleFileFound)
                lblColorRampName.Text = _selectedcolorrampname;
            switch (_displaymode)
            {
                case (1):
                    rbByLayer.Checked = true;
                    break;
                case (2):
                    rbByRiskLevel.Checked = true;
                    break;
                case (3):
                    rbColorRamp.Checked = true;
                    break;
            }
            cbRiskRankingChoices.SelectedIndex = _GISdata._cwdata.Configuration.BPRiskFieldtoUse;
            if ((_selectedcolorrampname.Length > 0) && (_styleFileFound))
            {

                ISymbologyStyleClass SC = axSymbologyControl1.GetStyleClass(esriSymbologyStyleClass.esriStyleClassColorRamps);
                int i;
                for (i = 0; i <= SC.ItemCount - 1;i++ )
                {
                    if (SC.GetItem(i).Name == _selectedcolorrampname)
                    {
                        SC.SelectItem(i);
                        break;
                    }
                }
                
                //IStyleGalleryItem serverStyleGalleryItem = axSymbologyControl1.GetStyleClass (axSymbologyControl1.StyleClass).GetSelectedItem();
                //System.Windows.Forms.MessageBox.Show(serverStyleGalleryItem.Name);
                
            }

            // Get Statistics to display
            LoadStatistics();
            // Done
            _displayModified = false;
            SetControls();
        }



        //-----------------------------------------------------------------------------------
        // Load the Risk Ranking Values
        //-----------------------------------------------------------------------------------
        private void LoadStatistics()
        {
            _GISdata._cwdata.linesectionGetStatistics(ref _totalLength, ref _unitlengths);
            // Clear Totals
            //........................................
            // Load Susceptibility
            //lvLengths.Items.Clear();
            //lvLengths.Columns.Clear();
            //lvLengths.Columns.Add("Unit");
            //lvLengths.Columns[0].Width = 50;
            //lvLengths.Columns[0].TextAlign = HorizontalAlignment.Center;
            //lvLengths.Columns.Add("Length", "Length");
            //lvLengths.Columns[1].Width = 110;
            //lvLengths.Columns[1].TextAlign = HorizontalAlignment.Right;
            lvLengths.Items.Clear();
            lvLengths.View = View.Details;
            int i;
            int c;
            string ss = "";
            ListViewItem lvi = null;
            for (i = 0; i <= (_unitlengths.Count - 1); i++)
            {
                lvi = new ListViewItem();
                ss = _unitlengths[i];
                c = Convert.ToInt32((ss.Substring(2, ss.Length - 2)));
                lvi.Text = ss.Substring(0, 1);                
                lvLengths.Items.Add(lvi);
                lvi.SubItems.Add(string.Format("{0:n0}",c));
            } // for
          
            // ALL (Total)
            lvi = new ListViewItem();
            lvi.Text = "ALL";
            lvLengths.Items.Add(lvi);
            lvi.SubItems.Add(string.Format("{0:n0}", _totalLength));
        }
      
        //.................................................
        //.................................................
        //.................................................
        private void SetControls()
        {
            axSymbologyControl1.Enabled = ((cbRiskRankingChoices.SelectedIndex >= 0) && (rbColorRamp.Checked) && (_styleFileFound));
            cbRiskRankingChoices.Enabled = ((rbColorRamp.Checked) || (rbByRiskLevel.Checked));
            btnOK.Enabled = (_wasmodified);
            if (_styleFileFound) 
                lblColorRampName.Text = SelectedColorRampName;
        }


        //.................................................
        //.................................................
        //.................................................
        private void rbByLayer_CheckedChanged(object sender, EventArgs e)
        {
            SetControls();
            _wasmodified = true;
            _displayModified = true;
            if (rbByRiskLevel.Checked)
            {
                _displaymode = 3;
            }                
            else if (rbColorRamp.Checked)
            {
                _displaymode = 2;
            }
            else
            {
                _displaymode = 1;
            }               
        }

        //.................................................
        //.................................................
        //.................................................
        private void axSymbologyControl1_OnItemSelected(object sender,
                                                        ESRI.ArcGIS.Controls.ISymbologyControlEvents_OnItemSelectedEvent e)
        {
            _wasmodified = true;
            _displayModified = true;
            lblColorRampName.Text = SelectedColorRampName;
            //Get the selected item
            m_styleGalleryItem = (IStyleGalleryItem) e.styleGalleryItem;
            _selectedcolorrampname = m_styleGalleryItem.Name;
        }


        private void GetRankingStatistics(string fieldName)
        {
            if (m_featureLayer == null) return;
            //Find the selected field in the feature layer
            IFeatureClass featureClass = m_featureLayer.FeatureClass;
            IField field = featureClass.Fields.get_Field(featureClass.FindField(fieldName));

            //Get a feature cursor
            ICursor cursor = (ICursor)m_featureLayer.Search(null, false);

            //Create a DataStatistics object and initialize properties
            IDataStatistics dataStatistics = new DataStatisticsClass();
            dataStatistics.Field = field.Name;
            dataStatistics.Cursor = cursor;

            //Get the result statistics
            IStatisticsResults statisticsResults = dataStatistics.Statistics;

            //Set the values min and max values
            _minriskvalue = statisticsResults.Minimum;
            _maxriskvalue = statisticsResults.Maximum;
        }


        //.................................................
        //.................................................
        //.................................................
        private void AssignColorRamp(string fieldname, List<RiskRankingRange> pRiskRankingRanges)
        {
            // Get Statistics Values for the Selected Field
            GetRankingStatistics(fieldname);
            if (_minriskvalue >= _maxriskvalue) return;
            if (pRiskRankingRanges.Count == 0) return;
            if (m_styleGalleryItem == null) return;
            if (m_featureLayer == null) return;
            m_classBreaksRenderer.Field = fieldname;
            m_classBreaksRenderer.BreakCount = pRiskRankingRanges.Count;
            m_classBreaksRenderer.MinimumBreak = 0;

            //Calculate the class interval by a simple mean value
            double interval = (_maxriskvalue - _minriskvalue) / pRiskRankingRanges.Count;

            //Get the color ramp
            IColorRamp colorRamp = (IColorRamp)m_styleGalleryItem.Item;
            //Set the size of the color ramp and recreate it
            colorRamp.Size = Convert.ToInt32(pRiskRankingRanges.Count);
            bool createRamp;
            colorRamp.CreateRamp(out createRamp);

            //Get the enumeration of colors from the color ramp
            IEnumColors enumColors = colorRamp.Colors;
            enumColors.Reset();
            double currentBreak = m_classBreaksRenderer.MinimumBreak;

            ISimpleLineSymbol simpleLineSymbol;
            //Loop through each class break
            for (int i = 0; i <= pRiskRankingRanges.Count - 1; i++)
            {
                //Set class break
                m_classBreaksRenderer.set_Break(i, currentBreak);
                //Create simple fill symbol and set color
                simpleLineSymbol = new SimpleLineSymbolClass();
                simpleLineSymbol.Color = enumColors.Next();
                //Add symbol to renderer
                m_classBreaksRenderer.set_Symbol(i, (ISymbol)simpleLineSymbol);
                currentBreak += interval;
            }
        }


        public static IRgbColor MakeRGBColor(byte R, byte G, byte B) 
        { 
         ESRI.ArcGIS.Display.RgbColor RgbClr = new ESRI.ArcGIS.Display.RgbColorClass(); 
         RgbClr.Red = R; 
         RgbClr.Green = G; 
         RgbClr.Blue = B; 
         return RgbClr; 
        }
        
        
        //.................................................
        //.................................................
        //.................................................
        private void AssignColorToRiskLevels(string fieldname, List<RiskRankingRange> pRiskRankingRanges)
        {
            // Get Statistics Values for the Selected Field
            GetRankingStatistics(fieldname);
            if (_minriskvalue >= _maxriskvalue) return;
            if (pRiskRankingRanges.Count == 0) return;
            if (m_featureLayer == null) return;
            m_classBreaksRenderer.Field = fieldname;
            m_classBreaksRenderer.BreakCount = pRiskRankingRanges.Count;
            m_classBreaksRenderer.MinimumBreak = 0;
            int i;
            double currentBreak = 0;
            ISimpleLineSymbol simpleLineSymbol;
            for (i = 0; i <= pRiskRankingRanges.Count - 1; i++)
            {
                currentBreak = pRiskRankingRanges[i].Upper;
                m_classBreaksRenderer.set_Break(i, currentBreak);
                //Create simple fill symbol and set color
                simpleLineSymbol = new SimpleLineSymbolClass();           
                simpleLineSymbol.Color = MakeRGBColor(pRiskRankingRanges[i].rColor, pRiskRankingRanges[i].gColor, pRiskRankingRanges[i].bColor);
                m_classBreaksRenderer.set_Symbol(i, (ISymbol)simpleLineSymbol);
            }
        }


        //.................................................
        //.................................................
        //.................................................
        private void AssignDefaultRenderer()
        {
            IRgbColor rgbColor = new RgbColorClass();
            rgbColor.Blue = 143;
            rgbColor.Red = 0;
            rgbColor.Green = 57;
            SimpleLineSymbolClass simpleLineSymbol = new SimpleLineSymbolClass();
            simpleLineSymbol.Color = rgbColor;
            m_SimpleRenderer.Symbol = (ISymbol)simpleLineSymbol;
        }

        //.................................................
        //.................................................
        //.................................................
        private void LoadColorRiskLevels()
        {
            string fieldname = "";
            m_classBreaksRenderer = new ClassBreaksRendererClass();
            switch (cbRiskRankingChoices.SelectedIndex)
            {
                case (0): // Overall
                    fieldname = "Overall_Ranking";
                    if (_GISdata._cwdata.RiskRankingRanges != null)
                        AssignColorToRiskLevels(fieldname, _GISdata._cwdata.RiskRankingRanges);
                    break;
                case (1): // Susceptibility
                    fieldname = "SusRanking";
                    if (_GISdata._cwdata.RiskRankingRangesSusc != null)
                        AssignColorToRiskLevels(fieldname, _GISdata._cwdata.RiskRankingRangesSusc);
                    break;
                case (2): // Consequence
                    fieldname = "ConRanking";
                    if (_GISdata._cwdata.RiskRankingRangesCons != null)
                        AssignColorToRiskLevels(fieldname, _GISdata._cwdata.RiskRankingRangesCons);
                    break;
            }
        }


        //.................................................
        //.................................................
        //.................................................
        private void LoadColorRamp()
        {
            string fieldname = "";
            //Create a new ClassBreaksRenderer and set properties
            if (rbByLayer.Checked)
            {
                AssignDefaultRenderer();
            }
            else
            {
                m_classBreaksRenderer = new ClassBreaksRendererClass();
                switch (cbRiskRankingChoices.SelectedIndex)
                {
                    case (0): // Overall
                        fieldname = "Overall_Ranking";
                        if (_GISdata._cwdata.RiskRankingRanges != null)
                            AssignColorRamp(fieldname, _GISdata._cwdata.RiskRankingRanges);
                        break;
                    case (1): // Susceptibility
                        fieldname = "SusRanking";
                        if (_GISdata._cwdata.RiskRankingRangesSusc != null)
                            AssignColorRamp(fieldname, _GISdata._cwdata.RiskRankingRangesSusc);
                        break;
                    case (2): // Consequence
                        fieldname = "ConRanking";
                        if (_GISdata._cwdata.RiskRankingRangesCons != null)
                            AssignColorRamp(fieldname, _GISdata._cwdata.RiskRankingRangesCons);
                        break;
                }
            }
        }


        //.................................................
        //.................................................
        //.................................................
        private void cbRiskRankingChoices_SelectedIndexChanged(object sender, EventArgs e)
        {
            _riskfieldtouse = cbRiskRankingChoices.SelectedIndex;
            _displayModified = true;
        }

        //.................................................
        //.................................................
        //.................................................
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (_displayModified)
            {
                if (rbByLayer.Checked)
                {
                    _layerDef.m_SimpleRenderer = m_SimpleRenderer;
                    _GISdata.SetLayerRendererBasicLine(_layerDef);
                    _displaymode = 1;
                    _layerDef.DisplayMode = 1;
                    _layerDef.DisplayByField = -1;
                    _GISdata._cwdata.Configuration.BPLSDisplayMode = _displaymode;
                    _GISdata._cwdata.Configuration.BPSavedColorRampName = "";
                    _GISdata._cwdata.Configuration.SaveConfigFile();
                }
                else
                    if (rbByRiskLevel.Checked)
                    {
                        LoadColorRiskLevels();
                        _layerDef.m_classBreaksRenderer = m_classBreaksRenderer;
                        _GISdata.SetLayerRenderer(_layerDef);
                        _displaymode = 2;
                        _layerDef.DisplayMode = 2;
                        _layerDef.DisplayByField = _riskfieldtouse;
                        _GISdata._cwdata.Configuration.BPLSDisplayMode = _displaymode;
                        _GISdata._cwdata.Configuration.BPSavedColorRampName = _selectedcolorrampname;
                        _GISdata._cwdata.Configuration.BPRiskFieldtoUse = _riskfieldtouse;
                        _layerDef.DisplayByField = _riskfieldtouse;
                        _GISdata._cwdata.Configuration.SaveConfigFile();
                    }                
                    else
                        if (rbColorRamp.Checked)
                        {
                            LoadColorRamp();
                            _layerDef.m_classBreaksRenderer = m_classBreaksRenderer;
                            _GISdata.SetLayerRenderer(_layerDef);
                            _displaymode = 3;
                            _layerDef.DisplayMode = 3;
                            _GISdata._cwdata.Configuration.BPLSDisplayMode = _displaymode;
                            _GISdata._cwdata.Configuration.BPSavedColorRampName = _selectedcolorrampname;
                            _GISdata._cwdata.Configuration.BPRiskFieldtoUse = _riskfieldtouse;
                            _layerDef.DisplayByField = _riskfieldtouse;
                            _GISdata._cwdata.Configuration.SaveConfigFile();                            
                        }
            }
        }


    }
}
