using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.Collections.Generic;

using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.ADF;
using ESRI.ArcGIS.SystemUI;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;


namespace COMPGIS1
{
  
    public sealed partial class MainForm : Form
    {
        #region class private members

        private IMapControl3 m_mapControl = null;
        private string m_mapDocumentName = string.Empty;
        private bool _closewhentaskcomplete = false;


        #endregion

        private ArcEngineClass _GISdata;
        private ArcEngineBP _BPGISdata;
        private ArcEngineSW _SWGISdata;

        private CompWorksDBConnect _cwdata;
        private List<string> _selectlayerfilters = new List<string>();
        private List<string> _appliedfilters = new List<string>();
        private CompWorksApps _cwapp = CompWorksApps.cwNone;

        // private properties
        private bool _adminlogin = false;

        private string _applicationname = "CompWorks {0} GIS Browser - [{1}]";
        private string _functioncode = "";
        private string _compworksdb = "";
        private string _systemfile = "";
        private string _featurekey = "";
        private string[] _passedparameters = new string[20];
        private int _passedparamcount = 0;
        private bool _changing_filter = false;
        private bool _settingMapLegend = false;

        public enum MapLegendMode { mlNone, mlBPLSRiskRanking }
        private MapLegendMode _maplegendmode = MapLegendMode.mlNone;
        private bool _maplegendvisible = false;

        private bool _apply_to_current_selection
        {
            get { return (cbApplyToSelection.Checked && cbApplyToSelection.Visible); }
            set { }
        }

        public bool AdminLogin
        {
            get { return _adminlogin; }
            set { }
        }

        public MapLegendMode LegendMode
        {
            get { return _maplegendmode; }
            set
            {
                SetMapLegendMode(value);
            }
        }

        public bool Maplegendvisible
        {
            get { return _maplegendvisible; }
            set { SetMapLegendVisibility(value); }
        }


        // Spatial Query Properties
        private double _lastusedspatialquerydistance = 20;
        private SelectedItemData _sqsourceitem;



        //--------------------------------------------------------------------------
        // MainForm Constructor
        //--------------------------------------------------------------------------

        #region class constructor

        public MainForm(string[] args)
        {
            InitializeComponent();
            ParseCommandLine(args);
            treeViewMain.Font = new Font(treeViewMain.Font, FontStyle.Bold);
        }

        #endregion

        //--------------------------------------------------------------------------
        // MainForm Load 
        //--------------------------------------------------------------------------
        private void MainForm_Load(object sender, EventArgs e)
        {
            
            // Check the Command Line Arguments
            if (CheckCmdLineArguments())
            {
                // Remove from TOC Tab Control
                tcTOC.TabPages.Remove(tabDisplay);
                tcMapLegend.TabPages.Remove(tabMapLegendRiskRanking);

                // Disable/Hide menu options that are to added based upon CompWorks app
                SetAppMenuOptions();

                // Set TOC
                tcTOC.SelectedTab = tabTOC;
                // get the MapControl
                m_mapControl = (IMapControl3) axMapControl1.Object;

                // make Data Connection GIS and Setup Form for the Selected CompWorks App
                switch (_cwapp)
                {
                    case (CompWorksApps.cwBuriedPiping):
                        _BPGISdata = new ArcEngineBP(_compworksdb,_systemfile);
                        if (_BPGISdata.InitializeResult != 0)
                        {
                            MessageBox.Show(_BPGISdata.InitializeError);
                            Application.Exit();
                        }

                        _GISdata = (ArcEngineClass)_BPGISdata;
                        tcMapLegend.TabPages.Add(tabMapLegendRiskRanking);
                        _applicationname = string.Format(_applicationname, "Buried Piping", _GISdata._cwdata.CompanyName);
                        break;

                    case (CompWorksApps.cwServiceWater):
                        _SWGISdata = new ArcEngineSW(_compworksdb,_systemfile);
                        if (_SWGISdata.InitializeResult != 0)
                        {
                            MessageBox.Show(_SWGISdata.InitializeError);
                            Application.Exit();
                        }

                        _GISdata = (ArcEngineClass)_SWGISdata;
                        _applicationname = string.Format(_applicationname, "Service Water", _GISdata._cwdata.CompanyName);
                        break;

                    case (CompWorksApps.cwFAC):
                        _applicationname = string.Format(_applicationname, "FAC", _GISdata._cwdata.CompanyName);
                        break;
                }

                //// disable the Save menu (since there is no document yet)
                menuSaveDoc.Enabled = false;
                //// Set reference pointer to CompWorks data class
                _cwdata = _GISdata._cwdata;
                if (File.Exists(_GISdata.GIS_Settings.GISmxdfile))
                {
                    axMapControl1.LoadMxFile(_GISdata.GIS_Settings.GISmxdfile);
                    _GISdata._activeview = axMapControl1.ActiveView;
                    _GISdata.LoadLayerObjectIDFields(_GISdata._workspace, _GISdata._layers);
                    SetSelectedMode();
                    axToolbarControl1.BackColor = SystemColors.Control;
                    InitFilterPanel();
                    InitSpatialSearchPanel();
                    InitContextMenus();
                    InitToolTips();
                    SetControls();
                    m_mapControl.Refresh();
                    InitMapLegends();
                }
                else
                {
                    if (_GISdata.GIS_Settings.GISmxdfile.Length == 0)
                        MessageBox.Show("The Map file to load has not been defined in Setup Utilities > GIS Settings");
                    else
                        MessageBox.Show("The specified Map file does not exist:\n\n" + _GISdata.GIS_Settings.GISmxdfile);
                    Application.Exit();
                }



                DoDatabaseUpdates();
            }
            else
            {
                Application.Exit();
            }
        }

        //--------------------------------------------------------------------------
        // Displays the menu options available for the specified CW Application
        //--------------------------------------------------------------------------
        private void SetAppMenuOptions()
        {
            // View menu
            lineSectionRiskRankingToolStripMenuItem.Visible = (_cwapp == CompWorksApps.cwBuriedPiping);
            mapLegendToolStripMenuItem.Visible = (_cwapp == CompWorksApps.cwBuriedPiping);
            viewToolStripMenuItem.Visible = (_cwapp == CompWorksApps.cwBuriedPiping);

            // Find menu
            findToolStripMenuItem.Visible = (_cwapp == CompWorksApps.cwBuriedPiping);
            linesBPToolStripMenuItem.Visible = (_cwapp == CompWorksApps.cwBuriedPiping);
            lineSectionsBPToolStripMenuItem.Visible = (_cwapp == CompWorksApps.cwBuriedPiping);
            boringBPExcavationLocationsToolStripMenuItem.Visible = (_cwapp == CompWorksApps.cwBuriedPiping);
            corrosionBPProbesToolStripMenuItem.Visible = (_cwapp == CompWorksApps.cwBuriedPiping);
        }

        //--------------------------------------------------------------------------
        // Checks to see if there are any pending database updates (i.e. Risk Ranking etc...)
        //--------------------------------------------------------------------------
        private void DoDatabaseUpdates()
        {

            switch (_cwapp)
            {
                case (CompWorksApps.cwBuriedPiping):
                    if (_BPGISdata.CheckForFieldUpdates())
                    {
                        BPDBUpdates updateform = new BPDBUpdates(_BPGISdata);
                        updateform.Show();
                        updateform.PerformUpdates();
                    }
                    break;

                case (CompWorksApps.cwServiceWater):
                    break;

                case (CompWorksApps.cwFAC):
                    break;
            }
        }

        //--------------------------------------------------------------------------
        // Checks the Command Line Arguments, primarily CWDatabase and CW System File
        // Returns false if they are not valid
        //--------------------------------------------------------------------------
        private bool CheckCmdLineArguments()
        {
            if (_cwapp == CompWorksApps.cwNone)
            {
                MessageBox.Show("The CompWorks application was not specified!\nPlease contact your System Administrator!");
                return false;
            }

            if (_compworksdb.Length == 0)
            {
                MessageBox.Show("The CompWorks database was not specified!\nPlease check Setup Utilities > GIS Settings");
                return false;
            }

            if (!File.Exists(_compworksdb))
            {
                MessageBox.Show(
                    "The specified CompWorks database does not exist!\nPlease check Setup Utilities > GIS Settings");
                return false;
            }

            if (_systemfile.Length == 0)
            {
                MessageBox.Show(
                    "The CompWorks workgroup database was not specified!\nPlease check Setup Utilities > GIS Settings");
                return false;
            }

            if (!File.Exists(_systemfile))
            {
                MessageBox.Show(
                    "The specified CompWorks workgroup database does not exist!\nPlease check Setup Utilities > GIS Settings");
                return false;
            }
            return true;
        }



        //--------------------------------------------------------------------------
        // SetSelectedMode
        //--------------------------------------------------------------------------
        private void SetSelectedMode()
        {
            string str = "";
            // Set to General mode
            Text = _applicationname;
//            Text = string.Format("CompWorks Buried Piping - [{0}]", _GISdata._cwdata.CompanyName);
            scMap.Panel2Collapsed = false;
            scSelections.Panel2Collapsed = true;

//            tcTOC.SelectedTab = tabSearch;
            menuStrip1.Visible = true;

            switch (_functioncode)
            {
                case (""): // Open Mode... 
                    {
                        break;
                    }


                // Functions

                case ("1001"):

                    //--------------------------------------------------------------------------------------------------------
                    // Buried Piping
                    //--------------------------------------------------------------------------------------------------------

                    // Select Line Sections within a specified radius of a Boring/Excavation Location and Associate with the B/E Location on OK
                    if (_cwapp != CompWorksApps.cwBuriedPiping) return;

                    {
                        if (_passedparamcount > 0)
                        {
                            if (_passedparameters[0].Length > 0)
                                tbRadius.Text = _passedparameters[0];
                            else
                                tbRadius.Text = "20";
                        }

                        // Set the filters for the selection set
                        menuStrip1.Visible = false;
                        _selectlayerfilters.Add(_GISdata._layers.LayerName(LayerSystemUse.suBPLineSections));
                        _selectlayerfilters.Add(_GISdata._layers.LayerName(LayerSystemUse.suBPBoringLocations));
                        spqry_BoringLocation_LineSections();
                        SelectionSetHasChanged();
                        ShowRootNode(LayerSystemUse.suBPLineSections);
                        Text = "CompWorks Buried Piping :: Select Line Sections near Boring/Excavation [" + _featurekey +
                               "]";
                        scMap.Panel2Collapsed = false;
                        scSelections.Panel2Collapsed = false;
                        tcCommandModes.SelectedTab = tabBE_LineSection_Association;
                        _closewhentaskcomplete = true;
                        break;
                    }

                case ("1002"): // SHOW SYSTEM - Select Lines and Line Sections of the specified System
                    {
                        if (_cwapp != CompWorksApps.cwBuriedPiping) return;
                        qry_System_Lines(_featurekey);
                        SelectionSetHasChanged();
                        break;
                    }

                case ("1003"): // SHOW Line
                    {
                        if (_cwapp != CompWorksApps.cwBuriedPiping) return;
                        qry_Lines(_featurekey, _GISdata._layers.LayerName(LayerSystemUse.suBPLines));
                        SelectionSetHasChanged();
                        break;
                    }

                case ("1004"): // SHOW Line Section
                    {
                        if (_cwapp != CompWorksApps.cwBuriedPiping) return;
                        selectLineSection(_featurekey, _GISdata._layers.LayerName(LayerSystemUse.suBPLineSections));
                        SelectionSetHasChanged();
                        break;
                    }

                case ("1005"): // SHOW Boring Locations
                    {
                        if (_cwapp != CompWorksApps.cwBuriedPiping) return;
                        selectBoringLocation(_featurekey, _GISdata._layers.LayerName(LayerSystemUse.suBPBoringLocations));
                        SelectionSetHasChanged();
                        break;
                    }

                case ("1006"): // SHOW CP Locations
                    {
                        if (_cwapp != CompWorksApps.cwBuriedPiping) return;
                        selectCPLocation(_featurekey, _GISdata._layers.LayerName(LayerSystemUse.suBPCorrosionProbes));
                        SelectionSetHasChanged();
                        break;
                    }

                //--------------------------------------------------------------------------------------------------------
                // Service Water
                //--------------------------------------------------------------------------------------------------------
                case ("2002"): // SHOW SYSTEM - Select Lines and Line Sections of the specified System
                    {
                        //if (_cwapp != CompWorksApps.cwBuriedPiping) return;
                        //qry_System_Lines(_featurekey);
                        SelectionSetHasChanged();
                        break;
                    }

                case ("2003"): // SHOW Line
                    {
                        //if (_cwapp != CompWorksApps.cwBuriedPiping) return;
                        //qry_Lines(_featurekey, _GISdata._layers.LayerName(LayerSystemUse.suBPLines));
                        SelectionSetHasChanged();
                        break;
                    }

                case ("2004"): // SHOW Line Section
                    {
                        //if (_cwapp != CompWorksApps.cwBuriedPiping) return;
                        //selectLineSection(_featurekey, _GISdata._layers.LayerName(LayerSystemUse.suBPLineSections));
                        SelectionSetHasChanged();
                        break;
                    }

                case ("2005"): // SHOW Component
                    {
                        //if (_cwapp != CompWorksApps.cwBuriedPiping) return;
                        //selectBoringLocation(_featurekey, _GISdata._layers.LayerName(LayerSystemUse.suBPBoringLocations));
                        SelectionSetHasChanged();
                        break;
                    }

            }
        }

        //--------------------------------------------------------------------------
        // Initialize the Map Legends 
        //--------------------------------------------------------------------------
        public void InitMapLegends()
        {
            // Risk Ranking Legend

            switch (_cwapp)
            {
                case (CompWorksApps.cwBuriedPiping):

                    gridRiskRankingLegend.Rows.Clear();
                    gridRiskRankingLegend.ColumnCount = 5;
                    gridRiskRankingLegend.Columns[0].Name = "Low";
                    gridRiskRankingLegend.Columns[1].Name = "Low-Medium";
                    gridRiskRankingLegend.Columns[2].Name = "Medium";
                    gridRiskRankingLegend.Columns[2].Name = "Medium-High";
                    gridRiskRankingLegend.Columns[2].Name = "High";
                    gridRiskRankingLegend.Rows[0].Height = 10;
                    cbMapLegendRiskFieldToUse.Items.Clear();
                    cbMapLegendRiskFieldToUse.DisplayMember = "ObjectName";
                    cbMapLegendRiskFieldToUse.ValueMember = "ObjectValue";

                    cbMapLegendRiskFieldToUse.Items.Add(new ListBoxNameValueStringObject
                    {
                        ObjectName = "Overall Ranking",
                        ObjectValue = "Overall_Ranking"
                    });
                    cbMapLegendRiskFieldToUse.Items.Add(new ListBoxNameValueStringObject
                    {
                        ObjectName = "Susceptibility Ranking",
                        ObjectValue = "SusRanking"
                    });
                    cbMapLegendRiskFieldToUse.Items.Add(new ListBoxNameValueStringObject
                    {
                        ObjectName = "Consequence Ranking",
                        ObjectValue = "ConRanking"
                    });
                    cbMapLegendRiskFieldToUse.SelectedIndex = _GISdata._cwdata.Configuration.BPRiskFieldtoUse;
                    break;

                case (CompWorksApps.cwServiceWater):
                    break;

                case (CompWorksApps.cwFAC):
                    break;
            } // switch

            LegendMode = MapLegendMode.mlNone;
            Maplegendvisible = false;
        }


        //--------------------------------------------------------------------------
        // Parse the Command Line Parameters
        //--------------------------------------------------------------------------
        public void ParseCommandLine(string[] args)
        {
            string str = "";
            int idx;
            for (idx = 0; idx <= (args.Length - 1); idx++)
            {
                string paramtype = "";
                str = args[idx];
                if (str.IndexOf("/") != 0) break;
                if (str.IndexOf("[") != 2) break;
                if (str.IndexOf("]") != (str.Length - 1)) break;
                paramtype = str.Substring(1, 1);

                switch (paramtype)
                {
                    case "a":
                        {
                            try
                            {
                                _cwapp = (CompWorksApps) Convert.ToInt32(str.Substring(3, str.Length - 4));
                            }
                            catch (Exception ex)
                            {
                                
                            }

                            break;
                        }

                    case "d":
                        {
                            _compworksdb = str.Substring(3, str.Length - 4);
                            break;
                        }
                    case "s": // system file
                        {
                            _systemfile = str.Substring(3, str.Length - 4);
                            break;
                        }
                    case "f": // function code
                        {
                            _functioncode = str.Substring(3, str.Length - 4);
                            break;
                        }
                    case "k": // feature key
                        {
                            _featurekey = str.Substring(3, str.Length - 4);
                            break;
                        }
                    case "p": // function parameter
                        {
                            _passedparamcount++;
                            _passedparameters[_passedparamcount - 1] = str.Substring(3, str.Length - 4);
                            break;
                        }

                } // switch
            } // for args...
        }


        #region Feature Execution Methods

        //--------------------------------------------------------------------------
        // spatial query LineSections near BoringLocation; function code 1001
        //--------------------------------------------------------------------------
        private void spqry_BoringLocation_LineSections()
        {
            // used in 1001
            //....................
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            int[] SpatialArray = new int[3000];
            string[] LSNames = new string[3000];
            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            if (tbRadius.Text.Length == 0) return;
            try
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                double radius = Convert.ToDouble(tbRadius.Text);
                if (_BPGISdata == null) return;
                _BPGISdata.ls_dospatialintersection(_featurekey, radius, ref pEnvelope, ref SpatialArray);
            }
            catch (Exception err)
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
            SetControls();
        }



        //--------------------------------------------------------------------------
        // Send data to CompWorks; used for function code 1001
        //--------------------------------------------------------------------------
        private bool AssociateLineSectionsToBoringLocation()
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return false;
            int counter = 0;
            int Zone_ID;
            try
            {
                while (counter <= (_GISdata._SelectedFeatures.Count - 1))
                {

                    if (_GISdata._SelectedFeatures[counter].layername ==
                        _GISdata._layers.LayerName(LayerSystemUse.suBPLineSections))
                    {
                        Zone_ID = Convert.ToInt32(_GISdata._SelectedFeatures[counter].CWKeyValue);
                        _GISdata._cwdata.linesection_SetAssociationToBoringLocation(Zone_ID, _featurekey);
                    }
                    counter++;
                }
            }
            catch (Exception err)
            {
                return false;
            }
            return true;
        }


        //--------------------------------------------------------------------------
        // Find Lines in System; function code 1002
        //--------------------------------------------------------------------------
        private int qry_System_Lines(string _featurekey)
        {
            // used in 1002
            //....................
            if (_cwapp != CompWorksApps.cwBuriedPiping) return 0;
            int numfound = 0;
            //string LSList = "";
            List<string> LineIDs = new List<string>();
            string system = _featurekey;

            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            try
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                if (_cwdata == null) return numfound;
                _cwdata.line_GetLinesInSystem(system, LineIDs);
                numfound = LineIDs.Count;
                if (LineIDs.Count > 0)
                {
                    _GISdata.line_SelectLinesInMap(LineIDs, ref pEnvelope,
                                                   _GISdata._layers.LayerName(LayerSystemUse.suBPLines));
                    pEnvelope.Expand(1.2, 1.2, true);
                    _GISdata._activeview.Extent = pEnvelope;
                    _GISdata._activeview.Refresh();
                    _GISdata._activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
                }
            }
            catch (Exception err)
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            SetControls();
            return numfound;
        }

        //--------------------------------------------------------------------------
        // Find Lines; function code 1003
        //--------------------------------------------------------------------------

        private void qry_Lines(string _featurekey, string _layerName)
        {
            // used in 1003
            //....................
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            int Feature_ID;
            string featname = "";
            string[] LineIDs = new string[50];
            string system = _featurekey;

            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            try
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                double radius = Convert.ToDouble(tbRadius.Text);
                //........................
                if (_cwdata == null) return;
                Feature_ID = Convert.ToInt32(_featurekey);
                featname = _GISdata._cwdata.line_GetLineName(Feature_ID);
                if (featname.Length > 0)
                {
                    string TLayer = _GISdata._layers.LayerName(LayerSystemUse.suBPLines);
                    LineIDs[0] = _featurekey;
                    _BPGISdata.line_SelectLines(LineIDs, ref pEnvelope, TLayer, featname);
                    pEnvelope.Expand(1.2, 1.2, true);
                    if (pEnvelope.XMax == pEnvelope.XMin) pEnvelope.XMax = pEnvelope.XMax + 1;
                    if (pEnvelope.YMax == pEnvelope.YMin) pEnvelope.YMax = pEnvelope.YMax + 1;
                    _GISdata._activeview.Extent = pEnvelope;
                    _GISdata._activeview.Refresh();
                }
                _GISdata._activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
            }
            catch (Exception err)
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            SetControls();
        }



        //--------------------------------------------------------------------------
        // Find Line Section; function code 1004
        //--------------------------------------------------------------------------
        private void selectLineSection(string _featurekey, string _layerName)
        {
            // used in 1004
            //....................
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            int Feature_ID;
            string featname = "";
            string[] LineIDs = new string[50];
            string system = _featurekey;

            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            try
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                //........................
                if (_cwdata == null) return;
                IActiveView pActiveView;
                pActiveView = axMapControl1.ActiveView;
                Feature_ID = Convert.ToInt32(_featurekey);
                featname = _GISdata._cwdata.linesection_GetLineSectionName(Feature_ID);
                if (featname.Length > 0)
                {
                    string TLayer = _layerName;
                    LineIDs[0] = _featurekey;
                    _BPGISdata.line_SelectLineSections(LineIDs, ref pEnvelope, TLayer, featname);
                    pEnvelope.Expand(1.2, 1.2, true);
                    if (pEnvelope.XMax == pEnvelope.XMin) pEnvelope.XMax = pEnvelope.XMax + 1;
                    if (pEnvelope.YMax == pEnvelope.YMin) pEnvelope.YMax = pEnvelope.YMax + 1;
                    pActiveView.Extent = pEnvelope;
                    pActiveView.Refresh();
                }
                pActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
            }
            catch (Exception err)
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            SetControls();
        }



        //--------------------------------------------------------------------------
        // Find Boring Location; function code 1005
        //--------------------------------------------------------------------------
        private void selectBoringLocation(string _featurekey, string _layerName)
        {
            // used in 1005
            //....................
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            string featname = "";
            string[] LineIDs = new string[50];
            string system = _featurekey;
            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            try
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                if (_cwdata == null) return;
                featname = _featurekey;
                string TLayer = _layerName;
                LineIDs[0] = _featurekey;
                _BPGISdata.SelectBoringLocation(LineIDs, ref pEnvelope, TLayer);
            }
            catch (Exception err)
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            SetControls();
        }


        //--------------------------------------------------------------------------
        // Find CP Location; function code 1006
        //--------------------------------------------------------------------------

        private void selectCPLocation(string _featurekey, string _layerName)
        {
            // used in 1006
            //....................
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            string featname = "";
            string[] LineIDs = new string[50];
            string system = _featurekey;

            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            try
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                if (_cwdata == null) return;
                IActiveView pActiveView;
                pActiveView = axMapControl1.ActiveView;
                featname = _featurekey;
                string TLayer = _layerName;
                LineIDs[0] = _featurekey;
                _BPGISdata.SelectCPLocation(LineIDs, ref pEnvelope, TLayer);
            }
            catch (Exception err)
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            SetControls();
        }



        //--------------------------------------------------------------------------
        // Remove the selected item from the Line Section list
        //--------------------------------------------------------------------------
        private void ls_Remove_Selection()
        {
            SetControls();
        }

        #endregion



        //--------------------------------------------------------------------------
        // Set Controls updates the controls when an item/value in the form has changed
        //--------------------------------------------------------------------------
        private void SetControls()
        {
            switch (_functioncode)
            {
                case ("1001"): // Select Line Sections within a specified radius of a Boring/Excavation Location
                    {
                        break;
                    }

                case ("1002"): // SHOW SYSTEM - Select Lines and Line Sections of the specified System
                    {
                        break;
                    }

                case ("1003"): // SHOW Line
                    {
                        break;
                    }

                case ("1004"): // SHOW Line Section
                    {
                        break;
                    }

                case ("1005"): // SHOW Boring Locations
                    {
                        break;
                    }

                case ("1006"): // SHOW CP Locations
                    {
                        break;
                    }
            }
            // Filter Panel Related
            btnFilterSave.Visible = false;
            if (cbActiveFilter.SelectedIndex >= 0)
            {
                // See if the selected filter is not saved; show the Save button
                SQLFilterClass filter =
                    ((ListBoxNameValueFilterObject) (cbActiveFilter.Items[cbActiveFilter.SelectedIndex])).ObjectValue;
                if (filter != null)
                    btnFilterSave.Visible = (filter.ID == 0);
            }
            cbActiveFilter.Enabled = (cbActiveFilter.Items.Count > 0);
            btnResetFilter.Enabled = (cbActiveFilter.SelectedIndex >= 0);
            btnRemoveFilter.Enabled = (cbActiveFilter.SelectedIndex >= 0);
            btnFilterEdit.Visible = ((cbActiveFilter.SelectedIndex >= 0) && (!btnFilterSave.Visible));
            btnExportTreeViewToExcel.Enabled = (treeViewMain.Nodes.Count > 0);
            // Display Panel on TOC Tabs
            if (rbByLayer.Checked)
            {
                
            }
            axSymbologyControl1.Enabled = (rbByRiskRankingColorRamp.Checked);
            cbRiskRankingChoices.Enabled = ((rbByRiskRankingColor.Checked) || (rbByRiskRankingColorRamp.Checked));
            btnApply.Enabled = (axSymbologyControl1.Enabled);

            // Legend related controls
            cbMapLegendRiskFieldToUse.Enabled = (rbMapLegendDisplayLSByRisk.Checked);
        }


        #region Form Opening Methods



        //--------------------------------------------------------------------------
        // Runs the Find Inspected Line Sections form
        //--------------------------------------------------------------------------
        private void RunFind_InspectedLineSections()
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            findLineSectionInspections runform = new findLineSectionInspections(_GISdata._cwdata._connectionstring);
            if (runform.ShowDialog() == DialogResult.OK)
            {
                if (runform.SelectedEvaluation)
                {
                    if ((runform.SelectedMethod.Length > 0) && (runform.SelectedWorkScope.Length > 0))
                    {
                        bool applytosel = runform.ApplyToCurrentSelection;
                        _BPGISdata.select_ClearAllSelections();
                        _BPGISdata.find_ShowEvaluatedLineSections(runform.SelectedMethod, runform.SelectedWorkScope,
                                                                applytosel);
                        string InspMethod = runform.SelectedMethod;
                        LoadTreeViewWithSelection(_GISdata._SelectedFeatures);
                        treeViewMain.ExpandAll();
                        DisplayQueryPanel("Evaluated Line Sections", "", "", "Type:", InspMethod,
                                          _GISdata._SelectedFeatures.Count, "", true);
                    }
                }
                else
                {
                    if (runform.SelectedSetupID > 0)
                    {
                        _BPGISdata.select_ClearAllSelections();
                        _BPGISdata.find_ShowInspectedLineSections(runform.SelectedSetupID);
                        string InspMethod = runform.SelectedMethod;
                        LoadTreeViewWithSelection(_GISdata._SelectedFeatures);
                        treeViewMain.ExpandAll();
                        DisplayQueryPanel("Inspected Line Sections", "", "", "Type:", InspMethod,
                                          _GISdata._SelectedFeatures.Count, "", true);
                    }
                }
            }
        }



        //--------------------------------------------------------------------------
        // Runs a spatial query Find looking for Target Layer Items Near Item in 
        //--------------------------------------------------------------------------
        private int Run_Spatial_Query()
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return -1;
            if (_BPGISdata == null) return -1;
            string SKey = "";
            string SLayer = "";
            string TLayer = "";
            double distance = 0;
            int[] SpatialArray = new int[3000];
            string[] LSNames = new string[3000];
            if (distance == 0) return -1;
            int retval = 0;
            _lastusedspatialquerydistance = distance;
            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            try
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                double radius = Convert.ToDouble(tbRadius.Text);
                retval = _BPGISdata.find_dospatialintersection(SKey, SLayer, TLayer, _lastusedspatialquerydistance,
                                                             ref pEnvelope,
                                                             ref SpatialArray);
            }
            catch (Exception err)
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }

            System.Windows.Forms.Cursor.Current = Cursors.Default;
            SetControls();
            return retval;
        }


        //--------------------------------------------------------------------------
        // Runs the Find Lines with Radiological Content
        //--------------------------------------------------------------------------
        private void RunFind_BP_Lines_With_Licensed_Material()
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            _BPGISdata.find_ShowLinesWithLicensedMaterial();
            LoadTreeViewWithSelection(_GISdata._SelectedFeatures);
            treeViewMain.ExpandAll();
            DisplayQueryPanel("Lines With Licensed Material", "", "", "", "", _GISdata._SelectedFeatures.Count, "", true);
        }


        //--------------------------------------------------------------------------
        // Runs the Find Line Sections Associated With LCOs
        //--------------------------------------------------------------------------
        private void RunFind_BP_Line_Sections_Associated_With_LCOs()
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            findLineSectionLCOForm runform = new findLineSectionLCOForm(_GISdata._cwdata._connectionstring);
            if (runform.ShowDialog() == DialogResult.OK)
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                int LCOValue = runform.LCO;
                string LCODesc = runform.LCODesc;
                bool applytocurrent = runform.ApplyToCurrentSelection;
                bool inclusivebelow = runform.InclusiveBelow;
                if (!applytocurrent)
                {
                    _BPGISdata.select_ClearAllSelections();
                }
                _BPGISdata.find_ShowLineSectionsWithLCO(LCOValue, inclusivebelow, applytocurrent);
                LoadTreeViewWithSelection(_GISdata._SelectedFeatures);
                treeViewMain.ExpandAll();
                DisplayQueryPanel("Line Sections Associated With LCOs", "LCO", LCODesc, "", "",
                                  _GISdata._SelectedFeatures.Count, "", true);
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
        }

        //--------------------------------------------------------------------------
        // Runs the Find Line Sections with GWUT Inspections by Level
        //--------------------------------------------------------------------------
        private void RunFind_BP_Line_Sections_With_GWUT_By_Level()
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            findGuidedWavebyLevel runform = new findGuidedWavebyLevel(_GISdata._cwdata._connectionstring);
            if (runform.ShowDialog() == DialogResult.OK)
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                int resultlevel = runform.SelectedLevel;
                _BPGISdata.select_ClearAllSelections();
                int SelLevel = runform.SelectedLevel;
                Boolean applytosel = runform.ApplyToCurrentSelection;
                _BPGISdata.find_ShowLineSectionsWithGWUTByLevel(SelLevel, applytosel);
                LoadTreeViewWithSelection(_GISdata._SelectedFeatures);
                treeViewMain.ExpandAll();
                DisplayQueryPanel("Line Sections With GWUT Inspections by Level", "Results Level", SelLevel.ToString(),
                                  "", "", _GISdata._SelectedFeatures.Count, "", true);
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
        }


        //--------------------------------------------------------------------------
        // Runs the Find Line Sections by selected risk ranking values
        //--------------------------------------------------------------------------
        private void RunFind_BP_LineSections_by_RiskRankingFactors()
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            findLineSectionsRiskRanking runform = new findLineSectionsRiskRanking(_GISdata._cwdata._connectionstring);
            if (runform.ShowDialog() == DialogResult.OK)
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                string sqlsusc = runform.SQLStatementSusc;
                string sqlcons = runform.SQLStatementCons;
                if ((sqlsusc.Length == 0) && (sqlcons.Length == 0)) return;
                _BPGISdata.select_ClearAllSelections();
                Boolean applytosel = runform.ApplyToCurrentSelection;
                _BPGISdata.find_ShowLineSectionsByRiskRanking(sqlsusc, sqlcons, applytosel);
                LoadTreeViewWithSelection(_GISdata._SelectedFeatures);
                treeViewMain.ExpandAll();
                DisplayQueryPanel("Line Sections by Risk Ranking", ":", "", "", "", _GISdata._SelectedFeatures.Count,
                                  runform.SelectDescription, false);
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
        }

        #endregion


        //--------------------------------------------------------------------------
        // Loads any Required Context Menu Items
        //--------------------------------------------------------------------------
        private void InitContextMenus()
        {
        }


        //--------------------------------------------------------------------------
        // Set the Tool Tips for the specified controls
        //--------------------------------------------------------------------------
        private void InitToolTips()
        {
            tooltipMain.SetToolTip(btnFilterSave, "Save the Current Filter");
            tooltipMain.SetToolTip(btnCreateFilter, "Create a New Custom Filter");
            tooltipMain.SetToolTip(btnResetFilter, "Clear the Current Filter");
            tooltipMain.SetToolTip(cbActiveFilter, "Select a Saved Filter");
            tooltipMain.SetToolTip(cbFilterLayerSelect, "Select the Layer to Filter");
            tooltipMain.SetToolTip(btnRemoveFilter, "Delete the Current Filter");
            tooltipMain.SetToolTip(btnExportTreeViewToExcel, "Export the Selected Features to Excel");
            tooltipMain.SetToolTip(cbApplyToSelection, "Apply the Filter to the Currently Selected Features");
            //tooltipMain.SetToolTip(cbSpatialSearchLayer, "Select the Layer to Search");            
        }


        //--------------------------------------------------------------------------
        // Display the Filter Panel
        //--------------------------------------------------------------------------
        private void DisplayFilterPanel()
        {
            tcTOC.SelectedTab = tabFilter;
        }


        //--------------------------------------------------------------------------
        // Sets up and display the Query Panel
        //--------------------------------------------------------------------------
        private void DisplayQueryPanel(string QueryDesc, string Param1lbl, string Param1Value, string Param2lbl,
                                       string Param2Value, int TotalCount, string SelectDescription,
                                       bool ShowParamLabels)
        {
            lblQueryDesc.Text = QueryDesc;
            lblQueryParamNote.Text = "";
            // Show Count
            if (TotalCount >= 0)
                lblQueryParamNote.Text = string.Format("{0}Features found: {1}\n", lblQueryParamNote.Text,
                                                       TotalCount.ToString());

            if (SelectDescription.Length > 0)
                lblQueryParamNote.Text = string.Format("{0}\n{1}\n", lblQueryParamNote.Text, SelectDescription);
            else
            {
                // Details label
                lblQueryParamNote.Text = string.Format("{0}\nDetails:\n", lblQueryParamNote.Text);

                // Parameter Values
                if ((Param1Value.Length > 0) && (Param1lbl.Length > 0))
                {
                    lblQueryParamNote.Text = string.Format("{0}\n{1}: {2}", lblQueryParamNote.Text, Param1lbl,
                                                           Param1Value);
                }
                if ((Param2Value.Length > 0) && (Param2lbl.Length > 0))
                {
                    lblQueryParamNote.Text = string.Format("{0}\n{1}: {2}", lblQueryParamNote.Text, Param2lbl,
                                                           Param2Value);
                }
            }
            tcTOC.SelectedTab = tabFind;
        }

        //--------------------------------------------------------------------------
        // Clear the Query Panel
        //--------------------------------------------------------------------------
        private void ClearQueryPanel()
        {
            lblQueryDesc.Text = "";
            lblQueryParamNote.Text = "";
            tcTOC.SelectedTab = tabTOC;
        }

        //--------------------------------------------------------------------------
        // Hide the Find panel
        //--------------------------------------------------------------------------
        private void HideQueryPanel()
        {
            //scSelections.Panel2Collapsed = true;
        }

        //--------------------------------------------------------------------------
        // Sets up the filter panel for use
        //--------------------------------------------------------------------------
        private void InitFilterPanel()
        {
            cbFilterLayerSelect.Items.Clear();
            cbActiveFilter.DisplayMember = "ObjectName";
            cbActiveFilter.ValueMember = "ObjectValue";
            int i;
            for (i = 0; i <= (_GISdata._layers.Count - 1); i++)
            {
                LayerSystemUse sysuse = _GISdata._layers.SystemUseOrdinal(i);
                if (sysuse > 0)
                {
                    cbFilterLayerSelect.Items.Add(_GISdata._layers.LayerName(sysuse));
                }
            }
        }

        //--------------------------------------------------------------------------
        // Sets up the spatial search panel for use
        //--------------------------------------------------------------------------
        private void InitSpatialSearchPanel()
        {
        }


        //--------------------------------------------------------------------------
        // Ensures that the specified root node is visible on the TreeVew
        //--------------------------------------------------------------------------
        private bool ShowRootNode(LayerSystemUse LayerSystemUse)
        {
            TreeNode[] treeNode;
            treeNode = treeViewMain.Nodes.Find(_GISdata._layers.LayerName(LayerSystemUse), true);
            if (treeNode.Length > 0)
            {
                treeNode[0].EnsureVisible();
                return true;
            }
            else
                return false;
        }


        //--------------------------------------------------------------------------
        // Load the tree view with the node list
        //--------------------------------------------------------------------------
        private void LoadTreeViewWithSelection(List<SelectedItemData> SelItems)
        {
            int idx;
            SelectedItemData selItem = new SelectedItemData();
            TreeNode[] treeNode;
            TreeNode newNode;
            TreeNode layerrootnode;
            treeViewMain.Nodes.Clear();
            for (idx = 0; idx <= (SelItems.Count - 1); idx++)
            {
                selItem = SelItems[idx];
                if (selItem != null)
                {
                    // See if the layer root Node exists
                    // if it doesn't, create it
                    // if it does, select it as the layerrootnode
                    treeNode = treeViewMain.Nodes.Find(selItem.layername, true);
                    if (treeNode.Length == 0)
                    {
                        layerrootnode = new TreeNode(selItem.layername);
                        layerrootnode.Name = selItem.layername;

                        // Add the metadata to the root node
                        SelectedRootNodeData rnd = new SelectedRootNodeData();
                        rnd.layerindex = _GISdata._layers.SystemUse(selItem.layername);
                        rnd.layername = selItem.layername;
                        rnd.activefilter = null;
                        layerrootnode.Tag = rnd;
                        treeViewMain.Nodes.Add(layerrootnode);

                        // Set the context menu for the root node
                        switch (_GISdata._layers.SystemUse(selItem.layername))
                        {
                            case (LayerSystemUse.suBPLines):
                                {
                                    layerrootnode.ContextMenuStrip = cmsBPRootLines;
                                    break;
                                }
                            case (LayerSystemUse.suBPLineSections):
                                {
                                    layerrootnode.ContextMenuStrip = cmsBPRootLineSections;
                                    break;
                                }
                            case (LayerSystemUse.suBPBoringLocations):
                                {
                                    layerrootnode.ContextMenuStrip = cmsBPRootBoringLocations;
                                    break;
                                }
                            case (LayerSystemUse.suBPCorrosionProbes):
                                {
                                    layerrootnode.ContextMenuStrip = cmsBPRootCorrosionProbes;
                                    break;
                                }

                            // Service Water
                            case (LayerSystemUse.suSWLines):
                                {
                                    layerrootnode.ContextMenuStrip = cmsSWRootLines;
                                    break;
                                }
                            case (LayerSystemUse.suSWFlowSegments):
                                {
                                    layerrootnode.ContextMenuStrip = cmsSWRootFlowSegments;
                                    break;
                                }
                            case (LayerSystemUse.suSWComponents):
                                {
                                    layerrootnode.ContextMenuStrip = cmsSWRootComponents;
                                    break;
                                }
                        }
                    }
                    else
                        layerrootnode = treeNode[0]; // Root Node Found
                    if (selItem.featuredescription != null)
                    {
                        if (selItem.featuredescription.Length > 0)
                        {
                            // Add the feature to the Layer Root Node
                            newNode = new TreeNode(selItem.featuredescription);
                            newNode.Name = selItem.layername;
                            newNode.Tag = selItem;
                            newNode.NodeFont = new Font(treeViewMain.Font, FontStyle.Regular);
                            newNode.SelectedImageIndex = 1;
                            layerrootnode.Nodes.Add(newNode);
                            // Set the context menu for the node
                            switch (_GISdata._layers.SystemUse(selItem.layername))
                            {
                                // BP Context Menus
                                case (LayerSystemUse.suBPLines):
                                    {
                                        newNode.ContextMenuStrip = cmsBPNodeLines;
                                        break;
                                    }
                                case (LayerSystemUse.suBPLineSections):
                                    {
                                        newNode.ContextMenuStrip = cmsBPNodeLineSection;
                                        break;
                                    }
                                case (LayerSystemUse.suBPBoringLocations):
                                    {
                                        newNode.ContextMenuStrip = cmsBPNodeBoringLocations;
                                        break;
                                    }
                                case (LayerSystemUse.suBPCorrosionProbes):
                                    {
                                        newNode.ContextMenuStrip = cmsBPNodeCorrosionProbes;
                                        break;
                                    }

                                // Service Water
                                case (LayerSystemUse.suSWLines):
                                    {
                                        newNode.ContextMenuStrip = cmsSWNodeLines;
                                        break;
                                    }
                                case (LayerSystemUse.suSWFlowSegments):
                                    {
                                        newNode.ContextMenuStrip = cmsSWNodeFlowSegments;
                                        break;
                                    }
                                case (LayerSystemUse.suSWComponents):
                                    {
                                        newNode.ContextMenuStrip = cmsSWNodeComponents;
                                        break;
                                    }

                            }
                        }
                    }
                }
            }
            SetRootLayerItemCountsInTitle();
        }


        //--------------------------------------------------------------------------
        // Adds the Record Count to the End of the Root Layer Name in brackets [xx]
        //--------------------------------------------------------------------------
        private void SetRootLayerItemCountsInTitle()
        {
            TreeNode[] treeNode;
            TreeNode newNode;
            int TotalSelection = 0;
            // Set the Root Nodes to display the number of records selected in brackets [xx]

            switch (_cwapp)
            {
                case (CompWorksApps.cwBuriedPiping):
                    treeNode = treeViewMain.Nodes.Find(_GISdata._layers.LayerName(LayerSystemUse.suBPLines), true);
                    if (treeNode.Length > 0)
                    {
                        newNode = treeNode[0];
                        newNode.Text = string.Format(newNode.Name + " [{0}]", newNode.Nodes.Count.ToString());
                        TotalSelection += newNode.Nodes.Count;
                    }
                    treeNode = treeViewMain.Nodes.Find(_GISdata._layers.LayerName(LayerSystemUse.suBPLineSections), true);
                    if (treeNode.Length > 0)
                    {
                        newNode = treeNode[0];
                        newNode.Text = string.Format(newNode.Name + " [{0}]", newNode.Nodes.Count.ToString());
                        TotalSelection += newNode.Nodes.Count;
                    }
                    treeNode = treeViewMain.Nodes.Find(_GISdata._layers.LayerName(LayerSystemUse.suBPBoringLocations), true);
                    if (treeNode.Length > 0)
                    {
                        newNode = treeNode[0];
                        newNode.Text = string.Format(newNode.Name + " [{0}]", newNode.Nodes.Count.ToString());
                        TotalSelection += newNode.Nodes.Count;
                    }
                    treeNode = treeViewMain.Nodes.Find(_GISdata._layers.LayerName(LayerSystemUse.suBPCorrosionProbes), true);
                    if (treeNode.Length > 0)
                    {
                        newNode = treeNode[0];
                        newNode.Text = string.Format(newNode.Name + " [{0}]", newNode.Nodes.Count.ToString());
                        TotalSelection += newNode.Nodes.Count;
                    }
                    break;
                case (CompWorksApps.cwServiceWater):
                    break;
                case (CompWorksApps.cwFAC):
                    break;
            }
            lblSelectedFeaturesTitle.Text = string.Format("Selected Features [{0}]", TotalSelection.ToString());
        }

        //--------------------------------------------------------------------------
        // Returns true if the specified layer is in the filter list
        //--------------------------------------------------------------------------
        private bool LayerInFilter(string layername)
        {
            if (layername.Length == 0) return false;
            if (_selectlayerfilters.Count == 0) return true;
            if (_selectlayerfilters.IndexOf(layername) >= 0)
                return true;
            else
                return false;
        }

        //--------------------------------------------------------------------------
        // Send the Selection Set to Excel
        //--------------------------------------------------------------------------
        private void ExportTreeViewToExcel()
        {
            // set a default file name
            saveExcelDialog.FileName = "SelectedFeatures.xls";
            // set filters - this can be done in properties as well
            saveExcelDialog.Filter = "Excel files (*.xls)|*.xls";
            if (saveExcelDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                    if (_GISdata.SendSelectedItemsToExcel(saveExcelDialog.FileName))
                    {
                        if (File.Exists(saveExcelDialog.FileName))
                        {
                            string pathstr = System.IO.Path.GetDirectoryName(saveExcelDialog.FileName);
                            System.Diagnostics.Process process = new System.Diagnostics.Process();
                            process.StartInfo.UseShellExecute = true;
                            process.StartInfo.FileName = @"explorer";
                            process.StartInfo.Arguments = string.Format("/select,\"{0}\"", saveExcelDialog.FileName);
                            process.Start();
                        }
                    }
                    else
                    {
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                        MessageBox.Show(string.Format("Unable to save the specified spreadsheet: [{0}] ",
                                                      saveExcelDialog.FileName));
                    }
                }
                finally
                {
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                }
            }
        }

        //--------------------------------------------------------------------------
        // The Selection Set in the GIS Viewer has changed... Load the List and Treeview
        //--------------------------------------------------------------------------
        private void SelectionSetHasChanged()
        {
            try
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                ESRI.ArcGIS.Geodatabase.IEnumFeature efeat =
                    axMapControl1.Map.FeatureSelection as IEnumFeature;
                (efeat as IEnumFeatureSetup).AllFields = true;
                efeat.Reset();
                _GISdata._SelectedFeatures.Clear();
                HideQueryPanel();
                ClearQueryPanel();
                IFeature feat;
                while ((feat = efeat.Next()) != null)
                {
                    SelectedItemData selItem = new SelectedItemData();
                    if ((_GISdata._layers.IsSystemLayer(((IFeatureClass) feat.Table).AliasName)) &&
                        (LayerInFilter(((IFeatureClass) feat.Table).AliasName)))
                    {
                        selItem.CWKeyValue = _GISdata.GetCWID_From_AttributeID(feat.OID,
                                                                               ((IFeatureClass) feat.Table).AliasName);
                        selItem.AttribObjectID = feat.OID;
                        selItem.layername = ((IFeatureClass) feat.Table).AliasName;
                        selItem.nodelevel = 0;
                        _GISdata._SelectedFeatures.Add(selItem);
                    }
                }
                // Load the descriptions for the selected features
                _GISdata.SetSelectionSetFeatureDescriptions();
                LoadTreeViewWithSelection(_GISdata._SelectedFeatures);
                treeViewMain.ExpandAll();
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
            }
            SetControls();
        }


        //--------------------------------------------------------------------------
        // Process a change to the treeView selection
        //--------------------------------------------------------------------------
        private void ProcessNodeSelection(TreeNode selNode)
        {
            if (selNode == null) return;
            try
            {
                switch (selNode.Level)
                {
                    case (0): // Layer Root Node - change form where required

                        // Change the filter panel if required
                        if (selNode.Tag != null)
                        {
                            SetFilterPanelLayer(((SelectedRootNodeData) selNode.Tag).layername);
                        }

                        break;
                    case (1): // Feature
                        SelectedItemData selItem = (SelectedItemData) selNode.Tag;
                        SetFilterPanelLayer(selItem.layername);
                        break;
                }
            }
            catch (Exception ex)
            {
            }
        }


        //--------------------------------------------------------------------------
        // Process a double-click on a selected tree node
        //--------------------------------------------------------------------------
        private void ProcessNodeDoubleClick(TreeNode selNode)
        {
            switch (selNode.Level)
            {                    
                case (0): // Layer Root Node
                    break;
                case (1): // Feature
                    SelectedItemData selItem = (SelectedItemData)selNode.Tag;

                    // make Data Connection GIS and Setup Form for the Selected CompWorks App
                    switch (_cwapp)
                    {
                        case (CompWorksApps.cwBuriedPiping):
                            _BPGISdata.ShowDetailForm(selItem);
                            break;

                        case (CompWorksApps.cwServiceWater):
                            _SWGISdata.ShowDetailForm(selItem);
                            break;

                        case (CompWorksApps.cwFAC):
                            break;
                    }
                    break;
            } // switch
        }


        //--------------------------------------------------------------------------
        // Opens the Synchronize Form
        //--------------------------------------------------------------------------
        private void adminFeatureConfiguration()
        {
            if (!adminGetLogin()) return;
            switch (_cwapp)
            {
                case (CompWorksApps.cwBuriedPiping):
                    BPSettingsForm sf = new BPSettingsForm(_BPGISdata);
                    sf.ShowDialog();
                    sf.Dispose();
                    break;

                case (CompWorksApps.cwServiceWater):
                    break;

                case (CompWorksApps.cwFAC):
                    break;
            }
        }

        //--------------------------------------------------------------------------
        // Opens the Layer Configuration Form
        //--------------------------------------------------------------------------
        private void adminLayerConfiguration()
        {
            if (!adminGetLogin()) return;
            LayerConfigForm sf = new LayerConfigForm(_GISdata);
            sf.ShowDialog();
            sf.Dispose();
        }

        //--------------------------------------------------------------------------
        // Opens the Database Update Form
        //--------------------------------------------------------------------------
        private void adminDoDatabaseUpdates()
        {
            if (!adminGetLogin()) return;


            switch (_cwapp)
            {
                case (CompWorksApps.cwBuriedPiping):
                    if (_BPGISdata.CheckForFieldUpdates())
                    {
                        BPDBUpdates updateform = new BPDBUpdates(_BPGISdata);
                        updateform.Show();
                        updateform.PerformUpdates();
                    }
                    else
                    {
                        MessageBox.Show("There are no updates required!");
                    }
                    break;

                case (CompWorksApps.cwServiceWater):
                    break;

                case (CompWorksApps.cwFAC):
                    break;
            }



        }

        #region Filter Related Methods

        //--------------------------------------------------------------------------
        // Set the layer displayed in the filter panel
        //--------------------------------------------------------------------------
        private void SetFilterPanelLayer(string layername)
        {
            if (cbFilterLayerSelect.SelectedIndex != cbFilterLayerSelect.Items.IndexOf(layername))
            {
                cbFilterLayerSelect.SelectedIndex = cbFilterLayerSelect.Items.IndexOf(layername);
            }
        }


        //--------------------------------------------------------------------------
        // If a filter was selected by a means other than selected the
        // filter dropdown (such as context menu)
        // This method is called to get the Filter Panel to match
        //--------------------------------------------------------------------------
        private void SetFilterPanelToMatchSelectedLayer(SQLFilterClass filter)
        {
            if (filter == null) return;
            bool _cfsave = _changing_filter;
            if (_changing_filter != true)
                _changing_filter = true;
            try
            {
                // Change the selected layer in the Panel to match the current filter
                // This also loads the Saved Filters for the layers
                SetFilteredLayer(_GISdata._layers.LayerName(filter.LayerIndex));
                // Select the Matching Filter
                int i;
                for (i = 0; i <= (cbActiveFilter.Items.Count - 1); i++)
                {
                    SQLFilterClass filtertocheck =
                        ((ListBoxNameValueFilterObject) (cbActiveFilter.Items[i])).ObjectValue;
                    if (filtertocheck.ID == filter.ID)
                    {
                        cbActiveFilter.SelectedIndex = i;
                        return;
                    }
                }
            }
            finally
            {
                _changing_filter = _cfsave;
            }
        }


        //--------------------------------------------------------------------------
        // Set the layer for the filter panel 
        // This is called when the selected layer in the treeview is changed
        // Not called when the selected filter for a layer has changed
        //--------------------------------------------------------------------------
        private void SetFilteredLayer(string layername)
        {
            // Load the Saved Filters
            SQLFilterClass activefilter = _GISdata._layers.ActiveFilter(layername);
            int ActiveFilterLayerIndex = -1;
            cbActiveFilter.Items.Clear();
            LayerDef sellayer = new LayerDef();
            _GISdata._layers.GetLayerDef(layername, ref sellayer);
            if (sellayer != null)
            {
                int i;
                for (i = 0; i <= (sellayer.SavedFilters.Count - 1); i++)
                {
                    SQLFilterClass filter = sellayer.GetSavedFilter(i);
                    if (filter != null)
                    {
                        cbActiveFilter.Items.Add(new ListBoxNameValueFilterObject
                            {
                                ObjectName = filter.FilterName,
                                ObjectValue = filter
                            });
                        // Get the index of the activefilter if it is specified
                        if (filter != null)
                        {
                            if (sellayer.ActiveFilter != null)
                                if (sellayer.ActiveFilter.ID == activefilter.ID)
                                    ActiveFilterLayerIndex = cbActiveFilter.Items.Count - 1;
                        }

                    }
                }
                // if an activefilter has been set for the layer, set the combo to display it
                if (ActiveFilterLayerIndex >= 0)
                    cbActiveFilter.SelectedIndex = ActiveFilterLayerIndex;
            }
        }


        //--------------------------------------------------------------------------
        // Load the specified ComboBox with the specified layer's saved filters
        //--------------------------------------------------------------------------
        private void LoadLayerFilterCombo(ref ComboBox filtercombo, string layername)
        {
            // Load the Saved Filters
            if (filtercombo == null) return;
            filtercombo.Items.Clear();
            LayerDef sellayer = new LayerDef();
            _GISdata._layers.GetLayerDef(layername, ref sellayer);
            if (sellayer != null)
            {
                int i;
                for (i = 0; i <= (sellayer.SavedFilters.Count - 1); i++)
                {
                    SQLFilterClass filter = sellayer.GetSavedFilter(i);
                    if (filter != null)
                    {
                        filtercombo.Items.Add(new ListBoxNameValueFilterObject
                            {
                                ObjectName = filter.FilterName,
                                ObjectValue = filter
                            });
                    }
                }
            }
        }


        //--------------------------------------------------------------------------
        // Create a filter for the selected layer
        //--------------------------------------------------------------------------
        private void CreateFilterForSelectedLayer()
        {
            if (cbFilterLayerSelect.SelectedIndex < 0) return; // no layer selected
            if (_changing_filter) return;
            _changing_filter = true;
            try
            {
                CreateAndExecuteLayerDBFilter(_GISdata._layers.SystemUse(cbFilterLayerSelect.Text));
//                SelectionSetHasChanged();
            }
            finally
            {
                _changing_filter = false;
            }
        }


        //--------------------------------------------------------------------------
        // Save the current filter if it has not been already
        //--------------------------------------------------------------------------
        private void SaveFilterForSelectedLayer()
        {
            DisplayFilterPanel();
            if (cbActiveFilter.SelectedIndex < 0) return; // no filter selected
            if (_changing_filter) return;
            _changing_filter = true;
            try
            {
                SQLFilterClass filter =
                    ((ListBoxNameValueFilterObject) (cbActiveFilter.Items[cbActiveFilter.SelectedIndex])).ObjectValue;
                if (filter != null)
                {
                    if (filter.ID == 0)
                    {
                        Filter_Edit feform = new Filter_Edit(filter, _GISdata._layers,
                                                             _GISdata._cwdata._connectionstring, true);
                        if (feform.ShowDialog() == DialogResult.OK)
                        {
                            if (_GISdata._cwdata.filter_AddNew(filter))
                            {
                                int idx = cbActiveFilter.SelectedIndex;
                                cbActiveFilter.Items.RemoveAt(idx);
                                cbActiveFilter.Items.Insert(0,
                                                            new ListBoxNameValueFilterObject
                                                                {
                                                                    ObjectName = filter.FilterName,
                                                                    ObjectValue = filter
                                                                });
                                idx = cbActiveFilter.Items.Count;
                                cbActiveFilter.SelectedIndex = 0;
                            }
                        }
                    }
                }
            }
            finally
            {
                _changing_filter = false;
            }
        }


        //--------------------------------------------------------------------------
        // Edit the current filter 
        //--------------------------------------------------------------------------
        private void EditFilterForSelectedLayer()
        {
            DisplayFilterPanel();
            if (cbActiveFilter.SelectedIndex < 0) return; // no filter selected
            if (_changing_filter) return;
            _changing_filter = true;
            try
            {
                SQLFilterClass filter =
                    ((ListBoxNameValueFilterObject) (cbActiveFilter.Items[cbActiveFilter.SelectedIndex])).ObjectValue;
                if (filter != null)
                {
                    if (filter.ID != 0)
                    {
                        Filter_Edit feform = new Filter_Edit(filter, _GISdata._layers,
                                                             _GISdata._cwdata._connectionstring, true);
                        if (feform.ShowDialog() == DialogResult.OK)
                        {
                            if (_GISdata._cwdata.filter_Update(filter))
                            {
                                int idx = cbActiveFilter.SelectedIndex;
                                cbActiveFilter.Items.RemoveAt(idx);
                                cbActiveFilter.Items.Add(new ListBoxNameValueFilterObject
                                    {
                                        ObjectName = filter.FilterName,
                                        ObjectValue = filter
                                    });
                                cbActiveFilter.SelectedIndex = (cbActiveFilter.Items.Count - 1);
                                ExecuteFilter(filter);
                            }
                        }
                    }
                }
            }
            finally
            {
                _changing_filter = false;
            }
        }

        //--------------------------------------------------------------------------
        // Set the active filter for a layer
        // This is called when the user selects an active filter
        //--------------------------------------------------------------------------
        private void SetActiveFilter(SQLFilterClass filter)
        {
            if (_changing_filter) return;
            _changing_filter = true;
            try
            {
                if (_apply_to_current_selection)
                {
                    // only run if the filter has not yet been applied
                    if (HasFilterBeenApplied(filter)) return;
                }
                else
                {
                    lbFiltersApplied.Items.Clear();
                }
                DisplayFilterPanel();
                if (filter == null) return;
                _GISdata._layers.ClearAllActiveFilters();
                if (ExecuteFilter(filter))
                {
                    try
                    {
                        AddFilterToApplied(filter);
                        LoadTreeViewWithSelection(_GISdata._SelectedFeatures);
                        treeViewMain.ExpandAll();
                    }
                    finally
                    {
                        System.Windows.Forms.Cursor.Current = Cursors.Default;
                    }
                    _GISdata._layers.SetActiveFilter(filter.LayerIndex, filter);
                }
            }
            finally
            {
                _changing_filter = false;
            }
        }

        //--------------------------------------------------------------------------
        // Returns true if the specified filter has been applied
        //--------------------------------------------------------------------------
        private bool HasFilterBeenApplied(SQLFilterClass filter)
        {
            if (filter == null) return false;
            return (_appliedfilters.IndexOf(filter.FilterName) >= 0);
        }

        //--------------------------------------------------------------------------
        // Adds the specified filter to the Applied List if it doesn't exist
        //--------------------------------------------------------------------------
        private void AddFilterToApplied(SQLFilterClass filter)
        {
            if (filter == null) return;
            if (lbFiltersApplied.Items.IndexOf(filter.FilterName) < 0)
            {
                _appliedfilters.Add(filter.FilterName);
                lbFiltersApplied.Items.Add(string.Format("{0}: {1}", lbFiltersApplied.Items.Count + 1, filter.FilterName));
                SetAppliedFilterVisible();
            }
        }

        //--------------------------------------------------------------------------
        // Clears the Applied Filter List
        //--------------------------------------------------------------------------
        private void ClearAppliedList()
        {
            _appliedfilters.Clear();
            lbFiltersApplied.Items.Clear();
            SetAppliedFilterVisible();
        }

        //--------------------------------------------------------------------------
        // Sets visiblity of the Applied Filter List
        //--------------------------------------------------------------------------
        private void SetAppliedFilterVisible()
        {
            lbFiltersApplied.Visible = (lbFiltersApplied.Items.Count > 0);
        }

        //--------------------------------------------------------------------------
        // Clears the filter for the selected layer
        //--------------------------------------------------------------------------
        private void ClearActiveFilter()
        {
            //DisplayFilterPanel();
            if (cbActiveFilter.SelectedIndex >= 0)
            {
                ClearAppliedList();
                SQLFilterClass filter =
                    ((ListBoxNameValueFilterObject) (cbActiveFilter.Items[cbActiveFilter.SelectedIndex])).ObjectValue;
                if (filter != null)
                {
                    _GISdata._layers.ClearActiveFilter(filter.LayerIndex);
                    _GISdata.ClearLayerSelectionSet(_GISdata._layers.LayerName(filter.LayerIndex));
                    treeViewMain.Nodes.Clear();
                }
                cbActiveFilter.SelectedIndex = -1;
                cbApplyToSelection.Checked = false;
                ClearAppliedList();
            }
        }


        //--------------------------------------------------------------------------
        // Deletes the filter for the selected layer
        //--------------------------------------------------------------------------
        private void RemoveActiveFilter()
        {
            DisplayFilterPanel();
            string ActiveFilterLayer = "";
            int FilterID = 0;
            if (cbActiveFilter.SelectedIndex >= 0)
            {
                SQLFilterClass filter =
                    ((ListBoxNameValueFilterObject) (cbActiveFilter.Items[cbActiveFilter.SelectedIndex])).ObjectValue;
                if (filter != null)
                {
                    DialogResult dialogResult =
                        MessageBox.Show(string.Format("Delete the Active Filter '{0}'?", filter.FilterName),
                                        "Delete Filter", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        ActiveFilterLayer = _GISdata._layers.LayerName(filter.LayerIndex);
                        _GISdata._layers.RemoveActiveFilter(filter.LayerIndex);
                        FilterID = filter.ID;
                        _GISdata._cwdata.filter_Remove(FilterID);
                        _GISdata.ClearLayerSelectionSet(_GISdata._layers.LayerName(filter.LayerIndex));
                        treeViewMain.Nodes.Clear();
                        SetFilteredLayer(ActiveFilterLayer);
                    }
                }
                cbActiveFilter.SelectedIndex = -1;
            }
        }


        //--------------------------------------------------------------------------
        // Creates and returns a new db filter for the specified layer
        //--------------------------------------------------------------------------
        private bool CreateLayerDBFilter(LayerSystemUse layerindex, ref SQLFilterClass newfilter)
        {
            DisplayFilterPanel();
            if (newfilter == null) return false;
            newfilter.LayerIndex = layerindex;
            Filter_Edit feform = new Filter_Edit(newfilter, _GISdata._layers, _GISdata._cwdata._connectionstring, false);
            if (feform.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                feform.Dispose();
                return true;
            }
            else
                return false;
        }

        //--------------------------------------------------------------------------
        // Creates then executes a filter for the specified layer
        //--------------------------------------------------------------------------
        private bool CreateAndExecuteLayerDBFilter(LayerSystemUse layerindex)
        {
            DisplayFilterPanel();
            SQLFilterClass newfilter = new SQLFilterClass();
            newfilter.LayerIndex = layerindex;
            newfilter.FilterName = "Custom Filter";
            newfilter.FilterDescription = "";
            if (CreateLayerDBFilter(layerindex, ref newfilter))
            {
                // Clear selection first if required
                if (!_apply_to_current_selection)
                    axMapControl1.Map.ClearSelection();
                // Execute Filter
                if (ExecuteFilter(newfilter))
                {
                    _GISdata._layers.AddFiltertoList(layerindex, newfilter);
                    cbActiveFilter.Items.Insert(0,
                                                new ListBoxNameValueFilterObject
                                                    {
                                                        ObjectName = newfilter.FilterName,
                                                        ObjectValue = newfilter
                                                    });
                    cbActiveFilter.SelectedIndex = 0;
                    LoadTreeViewWithSelection(_GISdata._SelectedFeatures);
                    treeViewMain.ExpandAll();
                    _GISdata._layers.SetActiveFilter(layerindex, newfilter);
                }
            }
            return false;
        }


        //--------------------------------------------------------------------------
        // Creates then executes a filter for the specified layer
        //--------------------------------------------------------------------------
        private bool ExecuteFilter(SQLFilterClass filter)
        {
            DisplayFilterPanel();
            if (filter == null) return false;
            if (filter.ConditionCount <= 0) return false;
            bool retvalue = false;
            try
            {
                System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                switch (_cwapp)
                {
                    case (CompWorksApps.cwBuriedPiping):
                        _BPGISdata.select_ClearAllSelections();
                        break;

                    case (CompWorksApps.cwServiceWater):
                        break;

                    case (CompWorksApps.cwFAC):
                        break;
                }
                retvalue = _GISdata.filter_ExecuteFilter(filter, (!_apply_to_current_selection));
                treeViewMain.ExpandAll();
            }
            finally
            {
                System.Windows.Forms.Cursor.Current = Cursors.Default;
                SetControls();
            }
            return retvalue;
        }

        #endregion



        #region Main Menu event handlers

        private void menuNewDoc_Click(object sender, EventArgs e)
        {
            //execute New Document command
            ICommand command = new CreateNewDocument();
            command.OnCreate(m_mapControl.Object);
            command.OnClick();
        }

        private void menuOpenDoc_Click(object sender, EventArgs e)
        {
            //execute Open Document command
            ICommand command = new ControlsOpenDocCommandClass();
            command.OnCreate(m_mapControl.Object);
            command.OnClick();
        }

        private void menuSaveDoc_Click(object sender, EventArgs e)
        {
            //execute Save Document command
            if (m_mapControl.CheckMxFile(m_mapDocumentName))
            {
                //create a new instance of a MapDocument
                IMapDocument mapDoc = new MapDocumentClass();
                mapDoc.Open(m_mapDocumentName, string.Empty);

                //Make sure that the MapDocument is not readonly
                if (mapDoc.get_IsReadOnly(m_mapDocumentName))
                {
                    MessageBox.Show("Map document is read only!");
                    mapDoc.Close();
                    return;
                }

                //Replace its contents with the current map
                mapDoc.ReplaceContents((IMxdContents) m_mapControl.Map);

                //save the MapDocument in order to persist it
                mapDoc.Save(mapDoc.UsesRelativePaths, false);

                //close the MapDocument
                mapDoc.Close();
            }
        }

        private void menuSaveAs_Click(object sender, EventArgs e)
        {
            //execute SaveAs Document command
            ICommand command = new ControlsSaveAsDocCommandClass();
            command.OnCreate(m_mapControl.Object);
            command.OnClick();
        }

        private void menuExitApp_Click(object sender, EventArgs e)
        {
            //exit the application
            Application.Exit();
        }

        #endregion

        //--------------------------------------------------------------------------
        //listen to MapReplaced event in order to update the statusbar and the Save menu
        //--------------------------------------------------------------------------
        private void axMapControl1_OnMapReplaced(object sender, IMapControlEvents2_OnMapReplacedEvent e)
        {
            //get the current document name from the MapControl
            m_mapDocumentName = m_mapControl.DocumentFilename;

            //if there is no MapDocument, diable the Save menu and clear the statusbar
            if (m_mapDocumentName == string.Empty)
            {
                menuSaveDoc.Enabled = false;
                tsFileName.Text = string.Empty;
            }
            else
            {
                //enable the Save manu and write the doc name to the statusbar
                menuSaveDoc.Enabled = true;
                tsFileName.Text = System.IO.Path.GetFileName(m_mapDocumentName);
            }
        }

        //--------------------------------------------------------------------------
        //--------------------------------------------------------------------------
        private void axMapControl1_OnMouseMove(object sender, IMapControlEvents2_OnMouseMoveEvent e)
        {
            statusBarXY.Text = string.Format("{0}, {1}  {2}", e.mapX.ToString("#######.##"),
                                             e.mapY.ToString("#######.##"),
                                             axMapControl1.MapUnits.ToString().Substring(4));
        }


        //--------------------------------------------------------------------------
        // Runs the Boring Location_LineSection Spatial Query for Line Section Association
        //--------------------------------------------------------------------------
        private void Reset_Click(object sender, EventArgs e)
        {
            spqry_BoringLocation_LineSections();
            SelectionSetHasChanged();
            ShowRootNode(LayerSystemUse.suBPLineSections); // zooms to be sure the Line Sections root node is showing
        }

        //--------------------------------------------------------------------------
        // Used in FunctionCode 1001
        //--------------------------------------------------------------------------
        private void AssociateLineSectionsToBoringLocations_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
            AssociateLineSectionsToBoringLocation();
            System.Windows.Forms.Cursor.Current = Cursors.Default;
            if (_closewhentaskcomplete)
                Application.Exit();
        }

        //--------------------------------------------------------------------------
        // Used in FunctionCode 1001
        //--------------------------------------------------------------------------
        private void AssociateLineSectionsToBoringLocationsCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //--------------------------------------------------------------------------
        // Used in FunctionCode 1001
        //--------------------------------------------------------------------------
        private void btnRemove_Click_1(object sender, EventArgs e)
        {
            ls_Remove_Selection();
        }


        //--------------------------------------------------------------------------
        //--------------------------------------------------------------------------
        private void lbResults_Click(object sender, EventArgs e)
        {
            SetControls();
        }

        //--------------------------------------------------------------------------
        //--------------------------------------------------------------------------
        private void axMapControl1_OnMouseDown(object sender, IMapControlEvents2_OnMouseDownEvent e)
        {

        }

        //--------------------------------------------------------------------------
        // The selection set of the Map Control has changed
        //--------------------------------------------------------------------------
        private void axMapControl1_OnSelectionChanged(object sender, EventArgs e)
        {
            if (!_changing_filter)
            {
                ClearActiveFilter();
            }
            SelectionSetHasChanged();
        }

        //--------------------------------------------------------------------------
        // Processes a double-click on the TreeView
        //--------------------------------------------------------------------------
        private void treeViewMain_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            ProcessNodeDoubleClick(e.Node);
        }


        #region TreeView Events for LineSection

        //--------------------------------------------------------------------------
        // Root Node Menu Opening
        //--------------------------------------------------------------------------
        private void cmsRootLineSections_Opening(object sender, CancelEventArgs e)
        {
            string layername = _GISdata._layers.LayerName(LayerSystemUse.suBPLineSections);
            LoadRootContextMenu(layername, cmsBPRootLineSections);
        }

        
        //--------------------------------------------------------------------------
        // Line Section Node Properties Option Click
        //--------------------------------------------------------------------------
        private void mnuLineSectionProperties_Click(object sender, EventArgs e)
        {
            if (treeViewMain.SelectedNode == null) return;
            if (treeViewMain.SelectedNode.Level != 1) return;
            SelectedItemData selItem = (SelectedItemData) treeViewMain.SelectedNode.Tag;
            if (selItem == null) return;
            if (selItem.layername != _GISdata._layers.LayerName(LayerSystemUse.suBPLineSections)) return;
            if (String.Compare(selItem.layername, _GISdata._layers.LayerName(LayerSystemUse.suBPLineSections)) == 0)
            {
                _BPGISdata.OpenForm_LineSectionDetails(Convert.ToInt32(selItem.CWKeyValue));
            }
        }

        //--------------------------------------------------------------------------
        // Line Section Node Inspections Options Click
        //--------------------------------------------------------------------------
        private void mnuLineSectionInspections(object sender, EventArgs e)
        {
            if (treeViewMain.SelectedNode == null) return;
            if (treeViewMain.SelectedNode.Level != 1) return;
            SelectedItemData selItem = (SelectedItemData) treeViewMain.SelectedNode.Tag;
            if (selItem == null) return;
            if (selItem.layername != _GISdata._layers.LayerName(LayerSystemUse.suBPLineSections)) return;
            if (String.Compare(selItem.layername, _GISdata._layers.LayerName(LayerSystemUse.suBPLineSections)) == 0)
            {
                _BPGISdata.OpenForm_LineSectionInspections(Convert.ToInt32(selItem.CWKeyValue));
            }

        }

        #endregion


        #region TreeView Events for Lines

        //--------------------------------------------------------------------------
        // Prepares the Root Node Menu for Opening
        //--------------------------------------------------------------------------
        private void LoadRootContextMenu(string layername, ContextMenuStrip cms)
        {
            if (layername.Length == 0) return;
            if (cms == null) return;

            // First set the Filter Menu to show Saved Filters for the layer
            ToolStripItem[] tsarray = null;
            ToolStripMenuItem filter_tsi = null;
            LayerSystemUse sysuse = _GISdata._layers.SystemUse(layername);
            switch (sysuse)
            {
                case (LayerSystemUse.suBPLines):
                    {
                        tsarray = cms.Items.Find("mnuFilterLines", true);
                        mnuCustomFilterSeparatorLines.Visible = false;
                        break;
                    }
                case (LayerSystemUse.suBPLineSections):
                    {
                        tsarray = cms.Items.Find("mnuFilterLineSections", true);
                        mnuCustomFilterSeparatorLineSections.Visible = false;
                        break;
                    }
                case (LayerSystemUse.suBPBoringLocations):
                    {
                        tsarray = cms.Items.Find("mnuFilterBoringLocations", true);
                        break;
                    }
                case (LayerSystemUse.suBPCorrosionProbes):
                    {
                        tsarray = cms.Items.Find("mnuFilterCorrosionProbes", true);
                        break;
                    }
            }
            if (tsarray.Length == 0) return;
            filter_tsi = (ToolStripMenuItem) tsarray[0];
            if (filter_tsi == null) return;
            filter_tsi.DropDownItems.Clear();
            filter_tsi.Enabled = false;

            // Custom filters are added to the menu dynamically
            // Load the LayerDef to access the saved filters 
            // Save the currently active filter for the layer (if there is one)
            SQLFilterClass activefilter = _GISdata._layers.ActiveFilter(layername);
            LayerDef sellayer = new LayerDef();
            _GISdata._layers.GetLayerDef(layername, ref sellayer);
            if (sellayer != null)
            {
                int i;
                for (i = 0; i <= (sellayer.SavedFilters.Count - 1); i++)
                {
                    SQLFilterClass filter = sellayer.GetSavedFilter(i);
                    if (filter != null)
                    {
                        ToolStripMenuItem mnuitem = new ToolStripMenuItem(filter.FilterName);
                        mnuitem.Tag = filter;
                        mnuitem.Click += mnuSavedFilterExecute_Click;
                        filter_tsi.DropDownItems.Add(mnuitem);
                        // If this is the active filter, Set it to Checked
                        if ((filter != null) && (sellayer.ActiveFilter != null))
                        {
                            mnuitem.Checked = (sellayer.ActiveFilter.ID == filter.ID);
                        }
                    }
                }
            }
            else
            {
                filter_tsi.Visible = false;
            }
            filter_tsi.Enabled = (filter_tsi.DropDownItems.Count > 0);
            mnuCustomFilterSeparatorLines.Visible = filter_tsi.Enabled;
            mnuCustomFilterSeparatorLineSections.Visible = filter_tsi.Enabled;
        }


        //--------------------------------------------------------------------------
        // Root Node Menu Opening Event
        //--------------------------------------------------------------------------
        private void cmsRootLines_Opening(object sender, CancelEventArgs e)
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            string layername = _GISdata._layers.LayerName(LayerSystemUse.suBPLines);
            LoadRootContextMenu(layername, cmsBPRootLines);
        }

        //--------------------------------------------------------------------------
        // Opening Line Properites Context Menu
        //--------------------------------------------------------------------------
        private void mnuLineProperties_Click(object sender, EventArgs e)
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            if (treeViewMain.SelectedNode == null) return;
            SelectedItemData selItem = (SelectedItemData) treeViewMain.SelectedNode.Tag;
            if (selItem == null) return;
            if (String.Compare(selItem.layername, _GISdata._layers.LayerName(LayerSystemUse.suBPLines)) == 0)
            {
                _BPGISdata.OpenForm_LineDetails(Convert.ToInt32(selItem.CWKeyValue));
            }
        }

        #endregion


        #region TreeView Events for Boring Locations

        //--------------------------------------------------------------------------
        // Root Node Menu Opening
        //--------------------------------------------------------------------------
        private void cmsRootBoringLocations_Opening(object sender, CancelEventArgs e)
        {
            // first look for custom filters and add to the list beneath the separator
        }

        //--------------------------------------------------------------------------
        // Boring Location Properties Option Click
        //--------------------------------------------------------------------------
        private void mnuBoringLocationProperties_Click(object sender, EventArgs e)
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            if (treeViewMain.SelectedNode == null) return;
            SelectedItemData selItem = (SelectedItemData) treeViewMain.SelectedNode.Tag;
            if (selItem == null) return;
            if (String.Compare(selItem.layername, _GISdata._layers.LayerName(LayerSystemUse.suBPBoringLocations)) == 0)
            {
                _BPGISdata.OpenForm_BoringLocationDetails(selItem.CWKeyValue);
            }
        }

        #endregion


        #region TreeView Events for Corrosion Probes

        //--------------------------------------------------------------------------
        // Root Node Menu Opening
        //--------------------------------------------------------------------------
        private void cmsRootCorrosionProbes_Opening(object sender, CancelEventArgs e)
        {
        }


        //--------------------------------------------------------------------------
        // Corrosion Probe Properties Option Click
        //--------------------------------------------------------------------------
        private void mnuCPProperties_Click(object sender, EventArgs e)
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            if (treeViewMain.SelectedNode == null) return;
            SelectedItemData selItem = (SelectedItemData) treeViewMain.SelectedNode.Tag;
            if (selItem == null) return;
            if (String.Compare(selItem.layername, _GISdata._layers.LayerName(LayerSystemUse.suBPCorrosionProbes)) == 0)
            {
                _BPGISdata.OpenForm_CorrosionProbeDetails(selItem.CWKeyValue);
            }
        }

        #endregion


        //--------------------------------------------------------------------------
        // Perform Admin Login if not already done
        //--------------------------------------------------------------------------
        private bool adminGetLogin()
        {
            if (AdminLogin)
                return true;
            else
            {
                LoginAdmin loginform = new LoginAdmin();
                DialogResult dr = loginform.ShowDialog();
                if (dr == DialogResult.Cancel)
                {
                    _adminlogin = false;
                    return false;
                }
                else
                {
                    if (loginform.LoginOK)
                    {
                        loginform.Dispose();
                        _adminlogin = true;
                        return true;
                    }
                    else
                    {
                        loginform.Dispose();
                        _adminlogin = false;
                        return false;
                    }
                }
            }
        }

        //--------------------------------------------------------------------------
        // Open a Layer's Settings Dialog
        //--------------------------------------------------------------------------
        private void OpenLayerSettings(string layerName)
        {
            LayerDef ld = new LayerDef();
            bool displaychanged = false;
            if (_GISdata._layers.GetLayerDef(layerName, ref ld))
            {
                _settingMapLegend = true;
                _GISdata._layers.OpenLayerSettings(ld.SystemUse, _GISdata, ref displaychanged);
                if (displaychanged)
                {
                    // Do something if needed
                    if ((ld.DisplayMode) > 1)
                    {
                        if (ld.DisplayByField >= 0)
                            cbMapLegendRiskFieldToUse.SelectedIndex = ld.DisplayByField;
                        else
                            cbMapLegendRiskFieldToUse.SelectedIndex = -1;

                        if (_maplegendmode != MapLegendMode.mlBPLSRiskRanking)
                            LegendMode = MapLegendMode.mlBPLSRiskRanking;
                        if (_maplegendmode == MapLegendMode.mlNone)
                            rbMapLegendDisplayLSByLayer.Checked = true;
                        else
                            rbMapLegendDisplayLSByRisk.Checked = true;
                        Maplegendvisible = true;
                    }

                    SetControls();
                }
                _settingMapLegend = false;
            }
        }


        //--------------------------------------------------------------------------
        // Set the active filter - 
        //--------------------------------------------------------------------------
        private void mnuSavedFilterExecute_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem menuoption = (ToolStripMenuItem) sender;
            if (menuoption.Tag != null)
            {
                SQLFilterClass filter = (SQLFilterClass) menuoption.Tag;
                SetActiveFilter(filter);
                SetFilterPanelToMatchSelectedLayer(filter);
            }

        }

        //--------------------------------------------------------------------------
        // Tree View - Expand All Nodes
        //--------------------------------------------------------------------------
        private void mnuExpandAll_Click(object sender, EventArgs e)
        {
            treeViewMain.ExpandAll();
        }

        //--------------------------------------------------------------------------
        // Tree View - Expand All Nodes
        //--------------------------------------------------------------------------
        private void mnuCollapseAll_Click(object sender, EventArgs e)
        {
            treeViewMain.CollapseAll();
        }


        //--------------------------------------------------------------------------
        // User changed the treeView selection
        //--------------------------------------------------------------------------
        private void treeViewMain_AfterSelect(object sender, TreeViewEventArgs e)
        {
            ProcessNodeSelection(e.Node);
            SetControls();
        }

        //--------------------------------------------------------------------------
        // The Selected Layer in filter panel has changed
        //--------------------------------------------------------------------------
        private void cbFilterLayerSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetFilteredLayer(cbFilterLayerSelect.Text);
            SetControls();
        }

        //--------------------------------------------------------------------------
        // The Selected Filter for a Layer has changed - user selected filter from dropdown
        //--------------------------------------------------------------------------
        private void cbActiveLayerFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_changing_filter)
            {
                if (cbActiveFilter.SelectedIndex >= 0)
                {
                    SQLFilterClass filter =
                        ((ListBoxNameValueFilterObject) (cbActiveFilter.Items[cbActiveFilter.SelectedIndex]))
                            .ObjectValue;
                    if (filter != null)
                    {
                        SetActiveFilter(filter);
                    }
                    SetControls();
                }
            }
        }

        //--------------------------------------------------------------------------
        // Create a filter for the selected layer
        //--------------------------------------------------------------------------
        private void btnCreateFilter_Click(object sender, EventArgs e)
        {
            CreateFilterForSelectedLayer();
            SetControls();
        }

        //--------------------------------------------------------------------------
        // Save the current filter if it has not been saved yet
        //--------------------------------------------------------------------------
        private void btnFilterSave_Click(object sender, EventArgs e)
        {
            SaveFilterForSelectedLayer();
            SetControls();
        }

        //--------------------------------------------------------------------------
        // Delete the current filter if it one is selected
        //--------------------------------------------------------------------------
        private void btnRemoveFilter_Click(object sender, EventArgs e)
        {
            RemoveActiveFilter();
            SetControls();
        }

        //--------------------------------------------------------------------------
        // Edit the current filter if it one is selected
        //--------------------------------------------------------------------------
        private void btnFilterEdit_Click(object sender, EventArgs e)
        {
            EditFilterForSelectedLayer();
            SetControls();
        }


        //--------------------------------------------------------------------------
        // Reset Active Layer Clicked
        //--------------------------------------------------------------------------
        private void btnResetFilter_Click(object sender, EventArgs e)
        {
            ClearActiveFilter();
            SetControls();
        }

        //--------------------------------------------------------------------------
        // Features Configuration
        //--------------------------------------------------------------------------
        private void synchronizeFeatureTablesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            adminFeatureConfiguration();
        }

        //--------------------------------------------------------------------------
        // Layer Configuration
        //--------------------------------------------------------------------------
        private void layerConfiguration_Click(object sender, EventArgs e)
        {
            adminLayerConfiguration();
        }

        //--------------------------------------------------------------------------
        // Do Database updates
        //--------------------------------------------------------------------------
        private void runDatabaseUpdatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            adminDoDatabaseUpdates();
        }

        //--------------------------------------------------------------------------
        // Find - Inspected Line Sections
        //--------------------------------------------------------------------------
        private void lineSectionsThatHaveBeenInspectedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RunFind_InspectedLineSections();
        }

        //--------------------------------------------------------------------------
        // Find - Lines with Licensed Material
        //--------------------------------------------------------------------------
        private void lineSectionsContainingLicensedMaterialradioactiveFluidsToolStripMenuItem1_Click(object sender,
                                                                                                     EventArgs e)
        {
            RunFind_BP_Lines_With_Licensed_Material();
        }

        //--------------------------------------------------------------------------
        // Find - Sections Associated With LCOs
        //--------------------------------------------------------------------------
        private void lineSectionsAssociatedWithLCOsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RunFind_BP_Line_Sections_Associated_With_LCOs();
        }

        //--------------------------------------------------------------------------
        // Find - Lines in a System
        //--------------------------------------------------------------------------
        private void linesInASystemToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_cwapp != CompWorksApps.cwBuriedPiping) return;
            SystemSelectionForm runform = new SystemSelectionForm(1, _GISdata._cwdata._connectionstring);
            if (runform.ShowDialog() == DialogResult.OK)
            {
                if (runform.SystemID > 0)
                {
                    _BPGISdata.select_ClearAllSelections();
                    int numfound = qry_System_Lines(runform.SystemID.ToString());
                    if (numfound == 0)
                    {
                        MessageBox.Show(string.Format("There were no Lines assigned to the System: '{0}'",
                                                      runform.SystemDesc));
                    }
                    else
                    {
                        SelectionSetHasChanged();
                        treeViewMain.ExpandAll();
                        DisplayQueryPanel("Lines in a System", "", "", "System:", runform.SystemDesc,
                                          _GISdata._SelectedFeatures.Count, "", true);
                    }
                }
                runform.Dispose();
            }
        }

        //--------------------------------------------------------------------------
        // Find - Spatial Query: Line Sections Near Boring Location
        //--------------------------------------------------------------------------
        private void lineSectionsNearbyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //RunFind_Spatial_Query_BoringLocation_LineSections();
        }


        //--------------------------------------------------------------------------
        // Export Selected Items to Excel
        //--------------------------------------------------------------------------
        private void btnExportTreeViewToExcel_Click(object sender, EventArgs e)
        {
            ExportTreeViewToExcel();
        }

        //--------------------------------------------------------------------------
        // Selected Spatial Query Target Layer has been selected / changed
        //--------------------------------------------------------------------------
        private void cbSQTLayer_SelectedIndexChanged(object sender, EventArgs e)
        {
            //LoadLayerFilterCombo(ref cbSQAddFilter, cbSQTLayer.Text);
            //SetControls();
        }

        //--------------------------------------------------------------------------
        // LineSections with Guided Wave Inspections by Level
        //--------------------------------------------------------------------------
        private void lineSectionsWithGuidedWaveInspectionsByLevelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RunFind_BP_Line_Sections_With_GWUT_By_Level();
            SetControls();
        }


        //--------------------------------------------------------------------------
        // Layer Selected on Spatial Search
        //--------------------------------------------------------------------------
        private void cbSpatialSearchLayer_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetControls();
        }


        //--------------------------------------------------------------------------
        // Search button on spatial search tab clicked
        //--------------------------------------------------------------------------
        private void btnSpatialSearch_Click(object sender, EventArgs e)
        {
            Run_Spatial_Query();
        }

        //--------------------------------------------------------------------------
        //--------------------------------------------------------------------------
        private void lineSectionsByRiskRankingValuesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RunFind_BP_LineSections_by_RiskRankingFactors();
        }

        //--------------------------------------------------------------------------
        //--------------------------------------------------------------------------
        private void axTOCControl1_OnDoubleClick(object sender, ITOCControlEvents_OnDoubleClickEvent e)
        {
            IBasicMap map = null;
            ILayer layer = null;
            System.Object other = null;
            System.Object index = null;
            esriTOCControlItem item = esriTOCControlItem.esriTOCControlItemNone;
            axTOCControl1.GetSelectedItem(ref item, ref map, ref layer, ref other, ref index);
            if (layer != null)
            {
                OpenLayerSettings(layer.Name);
            }
        }

        //--------------------------------------------------------------------------
        // Search button on spatial search tab clicked
        //--------------------------------------------------------------------------
        private void rbByLayer_CheckedChanged(object sender, EventArgs e)
        {
            SetControls();
        }

        //--------------------------------------------------------------------------
        // Map Legend / Risk Ranking tab Close  Button
        //--------------------------------------------------------------------------
        private void btnMapLegendRiskRankingClose_Click(object sender, EventArgs e)
        {
            Maplegendvisible = false;
        }

        //--------------------------------------------------------------------------
        // Set Map Legend Mode
        //--------------------------------------------------------------------------
        private void SetMapLegendMode(MapLegendMode _mapLegendMode)
        {
            switch (_mapLegendMode)
            {
                case (MapLegendMode.mlNone): // Off, closed
                    scMapLegend.Panel2Collapsed = true;
                    _maplegendmode = _mapLegendMode;
                    break;

                case (MapLegendMode.mlBPLSRiskRanking): // Risk Ranking Display Tab
                    if (_cwapp == CompWorksApps.cwBuriedPiping)
                    {
                        tcMapLegend.SelectedTab = tabMapLegendRiskRanking;
                        scMapLegend.Panel2Collapsed = (!_maplegendvisible);
                        _maplegendmode = _mapLegendMode;
                    }

                    if (_cwapp == CompWorksApps.cwServiceWater)
                    {
                    }

                    if (_cwapp == CompWorksApps.cwFAC)
                    {
                    }     
           
                    break;
            }
        }

        //--------------------------------------------------------------------------
        // Set Map Legend Visibility
        //--------------------------------------------------------------------------
        private void SetMapLegendVisibility(bool isVisible)
        {            
            switch (_maplegendmode)
            {
                case (MapLegendMode.mlNone): // Off, closed
                    scMapLegend.Panel2Collapsed = (!isVisible);
                    _maplegendvisible = isVisible;
                    break;

                case (MapLegendMode.mlBPLSRiskRanking): // Risk Ranking Display Tab
                    scMapLegend.Panel2Collapsed = (!isVisible);
                    _maplegendvisible = isVisible;
                    break;
            }
        }

        //--------------------------------------------------------------------------
        // Set Menu Items on Activation
        //--------------------------------------------------------------------------
        private void menuStrip1_MenuActivate(object sender, EventArgs e)
        {
            lineSectionRiskRankingToolStripMenuItem.Checked = ((_maplegendmode == MapLegendMode.mlBPLSRiskRanking) && (_maplegendvisible));
        }


        //--------------------------------------------------------------------------
        // Clicked on the Risk Ranking Legend View
        //--------------------------------------------------------------------------
        private void lineSectionRiskRankingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            lineSectionRiskRankingToolStripMenuItem.Checked = (!lineSectionRiskRankingToolStripMenuItem.Checked);
            Maplegendvisible = lineSectionRiskRankingToolStripMenuItem.Checked;
            SetControls();
        }

        //--------------------------------------------------------------------------
        // Clicked on Display by Risk Ranking
        //--------------------------------------------------------------------------
        private void rbMapLegendDisplayLSByRisk_CheckedChanged(object sender, EventArgs e)
        {
            if (_settingMapLegend) return;
            // Set the Mode for displaying the Line Section
            if (rbMapLegendDisplayLSByRisk.Checked)
            {
                if (cbMapLegendRiskFieldToUse.SelectedIndex >= 0)
                    try
                    {
                        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                        if (_BPGISdata.linesectionRenderByRiskRanking(
                            ((ListBoxNameValueStringObject)
                             (cbMapLegendRiskFieldToUse.Items[cbMapLegendRiskFieldToUse.SelectedIndex])).ObjectValue))
                        {
                            if (_maplegendmode != MapLegendMode.mlBPLSRiskRanking)
                                LegendMode = MapLegendMode.mlBPLSRiskRanking;
                            SetControls();
                        }
                    }
                    finally
                    {
                        System.Windows.Forms.Cursor.Current = Cursors.Default;                        
                    }
            }
        }

        //--------------------------------------------------------------------------
        // Clicked on Display by Layer Color
        //--------------------------------------------------------------------------
        private void rbMapLegendDisplayLSByLayer_CheckedChanged(object sender, EventArgs e)
        {
            if (_settingMapLegend) return;
            // Set the Mode for displaying the Line Section
            if (rbMapLegendDisplayLSByLayer.Checked)
            {
                try
                {
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                    _BPGISdata.linesection_AssignDefaultRenderer();
                    {
                        SetControls();
                    }
                }
                finally
                {
                    System.Windows.Forms.Cursor.Current = Cursors.Default;                    
                }
            }
        }

        //--------------------------------------------------------------------------
        // Clicked on Display by Risk Ranking - Which field to use
        //--------------------------------------------------------------------------
        private void cbMapLegendRiskFieldToUse_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_settingMapLegend) return;
            // Set the Mode for displaying the Line Section
            if (rbMapLegendDisplayLSByRisk.Checked)
            {
                if (cbMapLegendRiskFieldToUse.SelectedIndex >= 0)
                try
                {
                    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor;
                    if (_BPGISdata.linesectionRenderByRiskRanking(
                            ((ListBoxNameValueStringObject)
                             (cbMapLegendRiskFieldToUse.Items[cbMapLegendRiskFieldToUse.SelectedIndex])).ObjectValue))
                    {
                        if (_maplegendmode != MapLegendMode.mlBPLSRiskRanking)
                            LegendMode = MapLegendMode.mlBPLSRiskRanking;
                        SetControls();
                    }                 
                }
                finally
                {
                    System.Windows.Forms.Cursor.Current = Cursors.Default;
                }
            }
        }





    }


}


    