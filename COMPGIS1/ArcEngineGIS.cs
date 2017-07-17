using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.ADF;
using ESRI.ArcGIS.ADF.Connection.Local;
using ESRI.ArcGIS;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.SystemUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Display;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace COMPGIS1
{
    public class ArcEngineClass
    {
        // private properties
        // CompWorks Database and System Files
        private string _cwdatafile = "";
        private string _cwsystemfile = "";
        private int _initializeresult = 0;
        private string _initializeerror = "";
        private double _zoombuffersize = 1.2;

        // GIS Related Properties
        public GISDBSettings GIS_Settings;

        //...............................................
        // Selection Set Related
        //...............................................
        public List<SelectedItemData> _SelectedFeatures;

        //...............................................
        // Layer Related
        //...............................................
        // index in to the LayerDef SystemUse field
        public LayerDefs _layers;

        // public properties
        public bool IsInitialized = false;
        public CompWorksDBConnect _cwdata;
        public IWorkspace _workspace = null;
        public IActiveView _activeview = null;

        public int InitializeResult { get { return _initializeresult; } set {} }
        public string InitializeError { get { return _initializeerror; } set { } }



        //-----------------------------------------------------------------------------------------------
        // clear the selection set for a selected layer
        //-----------------------------------------------------------------------------------------------
        public void ClearLayerSelectionSet(string layername)
        {
            if (_activeview != null)
            {
                int index = GetIndexNumberFromLayerName(layername);
                if (index < 0) return;
                IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
                ((IFeatureSelection)featureLayer).Clear();
                _activeview.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);
                _activeview.Refresh();
                _activeview.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);
            }
        }


        //-----------------------------------------------------------------------------------------------
        // Clears the selection set for the specified layer(s)
        //-----------------------------------------------------------------------------------------------
        public void select_ClearLayerSelections(LayerSystemUse[] SystemUseToClear)
        {
            int i;
            for (i = 0; i <= (SystemUseToClear.Length - 1); i++)
            {
                string lname = _layers.LayerName(SystemUseToClear[i]);
                if (lname.Length > 0)
                {
                    ClearLayerSelectionSet(lname);
                }
            }
        }


        //-----------------------------------------------------------------------------------------------
        // Returns the list of layer names in the map
        //-----------------------------------------------------------------------------------------------
        public bool layer_getLayerList(ref List<string> layers)
        {
            if (_activeview != null)
            {
                IEnumLayer pEnumLayer = _activeview.FocusMap.Layers;
                pEnumLayer.Reset();
                ILayer pLayer = pEnumLayer.Next();
                while (pLayer != null)
                {
                    //if (pLayer.GetType() == typeof(ITable))
                    //{
                    //    layers.Add(pLayer.Name);
                    //}
                    layers.Add(pLayer.Name);
                    pLayer = pEnumLayer.Next();
                }
            }
            else
            {
                return false;
            }
            return (layers.Count > 0);
        }


        //-----------------------------------------------------------------------------------------------
        // Returns the Attribute Table CompWorks Index Field Name
        //   :: The value can be different since it is assigned by the system
        //   :: Typically the initial index field name will be OBJECTID
        //   :: Second would be OBJECTID1 etc...
        //-----------------------------------------------------------------------------------------------
        public string GetCWAttributeIndexFieldName(string layername)
        {
            string retval = "";
            if (layername.Length == 0) return retval;
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            retval = pTableOwner.OIDFieldName;
            return retval;
        }


        //-----------------------------------------------------------------------------------------------
        // Perform Radius Intersection of items in Target Layer (TLayer) around Object in SourceLayer (SLayer)
        //-----------------------------------------------------------------------------------------------
        protected void PerformRadiusIntersect(IWorkspace pWorkspace, double Radius, string SLayer, string TLayer, ref int[] SpatialArray, ref int ArrayCount, ref EnvelopeClass pEnvelope, string FeatureKey)
        {
            
            //Open FeatureClass
            IFeatureWorkspace pFeatureWorkspace = pWorkspace as IFeatureWorkspace;
            IFeatureClass pFeatureClass = pFeatureWorkspace.OpenFeatureClass(SLayer);
            IFeatureClass pIntersectedFeatureClass = pFeatureWorkspace.OpenFeatureClass(TLayer);
            //IFeature blockFeature = pBoringLocationsFeatureClass.GetFeature(SObjectID);
            IQueryFilter queryFilter1 = new QueryFilter { WhereClause = _layers.AttributeIndexName(SLayer) + " ='" + FeatureKey + "'" };
            
            //Code added to remove previous selections of lines or line sections
            int index = GetIndexNumberFromLayerName(TLayer);
            IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
            ((IFeatureSelection)featureLayer).Clear();

            IFeatureCursor featureCursor = pFeatureClass.Search(queryFilter1,true);
            IFeature blockFeature = featureCursor.NextFeature();
            IGeometry queryGeometry = blockFeature.ShapeCopy;
            // Buffer all the selected features by the buffer distance and create a new polygon element from each result
            ITopologicalOperator topologicalOperator;
            IElement element = null;
            IColor rgbColor = new RgbColorClass { Blue = 200 };
            ESRI.ArcGIS.Display.ISimpleLineSymbol simpleLineSymbol = new ESRI.ArcGIS.Display.SimpleLineSymbolClass();
            simpleLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;
            simpleLineSymbol.Color = rgbColor;
            simpleLineSymbol.Width = 4;

            topologicalOperator = (ESRI.ArcGIS.Geometry.ITopologicalOperator)blockFeature.Shape; // Explicit Cast

            element = new ESRI.ArcGIS.Carto.PolygonElementClass();
            element.Geometry = topologicalOperator.Buffer(Radius);
            IMap Map = _activeview.FocusMap;
            IGraphicsContainer graphicscontainer = Map as IGraphicsContainer;
            graphicscontainer.DeleteAllElements();
            //graphicscontainer.AddElement(element, 0);
            string layername = SLayer;
            SetBalloonPoint(FeatureKey, layername);
            //pActiveView.Selection.Clear();
            //((ISelectionSet)pActiveView.Selection).Select(queryFilter1,esriSelectionType.esriSelectionTypeHybrid,esriSelectionOption.esriSelectionOptionNormal,pWorkspace);
            CreateGraphicBuffersAroundSelectedFeatures(Radius);
            // Create the query filter.
            ISpatialFilter queryFilter = new SpatialFilterClass()
            {
                Geometry = element.Geometry,
                GeometryField = pIntersectedFeatureClass.ShapeFieldName,
                SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            };
            // Output
            //string objectID = SObjectID.ToString();
            int nameRes = pIntersectedFeatureClass.FindField(_layers.AttributeIndexName(TLayer));

            using (ComReleaser comReleaser = new ComReleaser())
            {
                IFeatureCursor cursor = pIntersectedFeatureClass.Search(queryFilter, true);
                comReleaser.ManageLifetime(cursor);
                IFeature feature = null;

                //// Set the filter to return only Line Sections.
                //int index = GetIndexNumberFromLayerName(pActiveView, "Line_Sections");
                //IFeatureLayer featureLayer = pActiveView.FocusMap.get_Layer(index) as IFeatureLayer;
                IFeatureSelection featureSelection = featureLayer as IFeatureSelection;
                featureSelection.Clear();
                featureSelection.SelectFeatures(queryFilter, ESRI.ArcGIS.Carto.esriSelectionResultEnum.esriSelectionResultNew, false);
                featureSelection.SelectionSymbol = simpleLineSymbol as ISymbol;
                //featureSelection.SelectionColor = rgbColor;
                _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);

                ArrayCount = 0;
                while ((feature = cursor.NextFeature()) != null)
                {
                    SpatialArray[ArrayCount] = Convert.ToInt32(feature.get_Value(nameRes));
                    ArrayCount++;
                    pEnvelope.Union(feature.Shape.Envelope);
                }
//                pEnvelope.Expand(1.2, 1.2, true);
                SetZoomEnvelope(ref pEnvelope);
                _activeview.Extent = pEnvelope;
                //pActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
           }
        }


        //-----------------------------------------------------------------------------------------------
        // Perform Radius Intersection of items in Target Layer (TLayer) around Object in SourceLayer (SLayer)
        //-----------------------------------------------------------------------------------------------
        protected void PerformRadiusIntersect(IWorkspace pWorkspace, double Radius, string SLayer, string TLayer, ref int[] SpatialArray, ref int ArrayCount, ref EnvelopeClass pEnvelope, string FeatureKey, bool ShowBalloon)
        {

            //Open FeatureClass
            IFeatureWorkspace pFeatureWorkspace = pWorkspace as IFeatureWorkspace;
            IFeatureClass pFeatureClass = pFeatureWorkspace.OpenFeatureClass(SLayer);
            IFeatureClass pIntersectedFeatureClass = pFeatureWorkspace.OpenFeatureClass(TLayer);
            //IFeature blockFeature = pBoringLocationsFeatureClass.GetFeature(SObjectID);
            IQueryFilter queryFilter1 = new QueryFilter { WhereClause = _layers.AttributeIndexName(SLayer) + " ='" + FeatureKey + "'" };

            //Code added to remove previous selections of lines or line sections
            int index = GetIndexNumberFromLayerName(TLayer);
            IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
            ((IFeatureSelection)featureLayer).Clear();

            IFeatureCursor featureCursor = pFeatureClass.Search(queryFilter1, true);
            IFeature blockFeature = featureCursor.NextFeature();
            IGeometry queryGeometry = blockFeature.ShapeCopy;
            // Buffer all the selected features by the buffer distance and create a new polygon element from each result
            ITopologicalOperator topologicalOperator;
            IElement element = null;
            IColor rgbColor = new RgbColorClass { Blue = 200 };
            ESRI.ArcGIS.Display.ISimpleLineSymbol simpleLineSymbol = new ESRI.ArcGIS.Display.SimpleLineSymbolClass();
            simpleLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;
            simpleLineSymbol.Color = rgbColor;
            simpleLineSymbol.Width = 4;

            topologicalOperator = (ESRI.ArcGIS.Geometry.ITopologicalOperator)blockFeature.Shape; // Explicit Cast

            element = new ESRI.ArcGIS.Carto.PolygonElementClass();
            element.Geometry = topologicalOperator.Buffer(Radius);
            IMap Map = _activeview.FocusMap;
            IGraphicsContainer graphicscontainer = Map as IGraphicsContainer;
            graphicscontainer.DeleteAllElements();
            //graphicscontainer.AddElement(element, 0);
            string layername = SLayer;
            if (ShowBalloon)
                SetBalloonPoint(FeatureKey, layername);
            //pActiveView.Selection.Clear();
            //((ISelectionSet)pActiveView.Selection).Select(queryFilter1,esriSelectionType.esriSelectionTypeHybrid,esriSelectionOption.esriSelectionOptionNormal,pWorkspace);
            CreateGraphicBuffersAroundSelectedFeatures(Radius);
            // Create the query filter.
            ISpatialFilter queryFilter = new SpatialFilterClass()
            {
                Geometry = element.Geometry,
                GeometryField = pIntersectedFeatureClass.ShapeFieldName,
                SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            };
            // Output
            //string objectID = SObjectID.ToString();
            int nameRes = pIntersectedFeatureClass.FindField(_layers.AttributeIndexName(TLayer));

            using (ComReleaser comReleaser = new ComReleaser())
            {
                IFeatureCursor cursor = pIntersectedFeatureClass.Search(queryFilter, true);
                comReleaser.ManageLifetime(cursor);
                IFeature feature = null;

                //// Set the filter to return only Line Sections.
                //int index = GetIndexNumberFromLayerName(pActiveView, "Line_Sections");
                //IFeatureLayer featureLayer = pActiveView.FocusMap.get_Layer(index) as IFeatureLayer;
                IFeatureSelection featureSelection = featureLayer as IFeatureSelection;
                featureSelection.Clear();
                featureSelection.SelectFeatures(queryFilter, ESRI.ArcGIS.Carto.esriSelectionResultEnum.esriSelectionResultNew, false);
                featureSelection.SelectionSymbol = simpleLineSymbol as ISymbol;
                //featureSelection.SelectionColor = rgbColor;
                _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);

                ArrayCount = 0;
                while ((feature = cursor.NextFeature()) != null)
                {
                    SpatialArray[ArrayCount] = Convert.ToInt32(feature.get_Value(nameRes));
                    ArrayCount++;
                    pEnvelope.Union(feature.Shape.Envelope);
                }
                //                pEnvelope.Expand(1.2, 1.2, true);
                SetZoomEnvelope(ref pEnvelope);
                _activeview.Extent = pEnvelope;
                //pActiveView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
            }
        }


        //-----------------------------------------------------------------------------------------------
        // Performs an Inner Join on the two lists; Returns the resulting set in the ResultList parameter
        //-----------------------------------------------------------------------------------------------
        public bool joinlists_Inner(List<SelectedItemData> L1, List<SelectedItemData> L2,
                                    ref List<SelectedItemData> ResultList)
        {
            if (ResultList == null) return false;
            if (L1 == null) return false;
            if (L2 == null) return false;
            int i;
            SelectedItemData sid;
            for (i = 0; i <= (L1.Count - 1); i++)
            {
                sid = L2.Find(x => x.CWKeyValue == L1[i].CWKeyValue); 
                if (sid != null)
                {
                    ResultList.Add(sid);
                }
            }
            return true;
        }

        //-----------------------------------------------------------------------------------------------
        // Assigns the values from one SelectedItemData list to another
        //-----------------------------------------------------------------------------------------------
        public bool lists_Assign(List<SelectedItemData> L1, ref List<SelectedItemData> ResultList)
        {
            if (ResultList == null) return false;
            ResultList.Clear();
            if (L1 == null) return false;
            int i;
            for (i = 0; i <= (L1.Count - 1); i++)
            {
                ResultList.Add(L1[i]);
            }
            return true;
        }


        //-----------------------------------------------------------------------------------------------
        // Executes the filter and populates the _selectionset with the results
        //-----------------------------------------------------------------------------------------------
        public bool filter_ExecuteFilter(SQLFilterClass filter, bool ClearSelectedFeatures)
        {
            if (filter == null) return false;
            if (filter.ConditionCount <= 0) return false;
            string TLayer = _layers.LayerName(filter.LayerIndex);
            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            List<SelectedItemData> tempList = new List<SelectedItemData>();
            List<SelectedItemData> resultList = new List<SelectedItemData>();
            ESRI.ArcGIS.Carto.IMap map = _activeview.FocusMap;
            int i;

            if (_cwdata.ExecuteFiltertoSelectionSet(filter, _layers, ref tempList) == true)
            {
                if (ClearSelectedFeatures)
                {
                    _SelectedFeatures.Clear();
                    for (i = 0; i <= (tempList.Count - 1); i++)
                        _SelectedFeatures.Add(tempList[i]);
                }
                else // Combining with Existing Selection;  Perform an Inner Join on the Lists and filter for layer
                {
                    joinlists_Inner(tempList, _SelectedFeatures, ref resultList);
                    _SelectedFeatures.Clear();
                    map.ClearSelection();
                    for (i = 0; i <= (resultList.Count - 1); i++)
                    {
                        if (resultList[i].layername == (_layers.LayerName((LayerSystemUse)filter.LayerIndex)))
                            _SelectedFeatures.Add(resultList[i]);
                    }
                        
                }
                if (select_SelectFeaturesInMap(ref _SelectedFeatures, ref pEnvelope, TLayer) > 0)
                {
                    SetZoomEnvelope(ref pEnvelope);
                    _activeview.Extent = pEnvelope;
                    _activeview.Refresh();
                    _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
                    return true;
                }
            }
            return false;
        }


        //-----------------------------------------------------------------------------------------------
        // Returns an Envelope Zoomed according to dimensions and buffers when height or width is 0
        //-----------------------------------------------------------------------------------------------
        protected void SetZoomEnvelope(ref ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope)
        {
            pEnvelope.XMax = pEnvelope.XMax + 5;
            pEnvelope.XMin = pEnvelope.XMin - 5;
            pEnvelope.YMax = pEnvelope.YMax + 5;
            pEnvelope.YMin = pEnvelope.YMin - 5;
            if (pEnvelope.XMax == pEnvelope.XMin) pEnvelope.XMax = pEnvelope.XMax + .00001;
            if (pEnvelope.YMax == pEnvelope.YMin) pEnvelope.YMax = pEnvelope.YMax + .00001;
            pEnvelope.Expand(_zoombuffersize, _zoombuffersize, true);
        }

        //-----------------------------------------------------------------------------------------------
        // Select the Features from the SelectedFeatures list and pass back the envelope of the new selection
        // Overload 1
        //-----------------------------------------------------------------------------------------------
        public int select_SelectFeaturesInMap(ref List<SelectedItemData> SelFeatures, ref EnvelopeClass pEnvelope, string TLayer)
        {
            int retcount = 0;
            IColor rgbColor = new RgbColorClass { Red = 100 };
            ESRI.ArcGIS.Display.ISimpleLineSymbol simpleLineSymbol = new ESRI.ArcGIS.Display.SimpleLineSymbolClass();
            simpleLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;
            simpleLineSymbol.Color = rgbColor;
            simpleLineSymbol.Width = 4;
            IWorkspace pLWorkspace;
            pLWorkspace = FileGDBWorkspaceFromPath(GIS_Settings.GISRootFolder);
            IFeatureWorkspace pFeatureWorkspace = pLWorkspace as IFeatureWorkspace;
            IFeatureClass pLineFeatureClass = pFeatureWorkspace.OpenFeatureClass(TLayer);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            int nameRes = pLineFeatureClass.FindField(_layers.AttributeIndexName(TLayer));

            // Clear the Selected Items List (if needed) and add the newly selected items

            using (ComReleaser comReleaser = new ComReleaser())
            {
                //IFeature feature = null;
                int index = GetIndexNumberFromLayerName(TLayer);
                const int maxConditionsPerStatement = 10;
                IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;

                IFeatureSelection featureSelection = featureLayer as IFeatureSelection;
                int counter = 0;
                int batchcounter = 0;
                string logicaloperator = "";
                string whereclause = "(";
                string attindexname = _layers.AttributeObjectIndexName(TLayer);
                string dbindexname = _layers.CWIndexName(TLayer);

                // Build the Where Clause to search with using all the keys and OR'ing them
                while (counter <= (SelFeatures.Count - 1))
                {
                    if ((batchcounter > 0) && (counter <= (SelFeatures.Count - 1)))
                        logicaloperator = " OR "; // Only use when not the first or last condition 
                    whereclause += string.Format("{0} ({1} = {2})", logicaloperator, attindexname, SelFeatures[counter].AttribObjectID);
                    if ((batchcounter > maxConditionsPerStatement) || (counter == (SelFeatures.Count - 1))) // if maxstatements OR last item
                    {
                        whereclause += ")";
                        // Set the filter
                        queryFilter.WhereClause = whereclause;
                        IFeatureCursor cursor = pLineFeatureClass.Search(queryFilter, true);
                        comReleaser.ManageLifetime(cursor);
                        IFeature pFeature = cursor.NextFeature();
                        while (pFeature != null)
                        {
                            featureSelection.Add(pFeature);
                            pEnvelope.Union(pFeature.Shape.Envelope);
                            retcount++;
                            pFeature = cursor.NextFeature();
                        }
                        // Clear the values for the next batch
                        logicaloperator = "";
                        whereclause = "(";
                        batchcounter = 0;
                    }
                    else
                        batchcounter++;
                    counter++;
                } // while

                featureSelection.SelectionSymbol = simpleLineSymbol as ISymbol;
                _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
            }
            return retcount;
        }


        //-----------------------------------------------------------------------------------------------
        // Select the Features with KeyIDs and pass back the envelope of the new selection
        // Overload 2
        //-----------------------------------------------------------------------------------------------
        public int select_SelectFeaturesInMap(List<string> KeyIDs, ref EnvelopeClass pEnvelope, string TLayer, bool ClearExisting)
        {
            int retcount = 0;
            IColor rgbColor = new RgbColorClass { Red = 100 };
            ESRI.ArcGIS.Display.ISimpleLineSymbol simpleLineSymbol = new ESRI.ArcGIS.Display.SimpleLineSymbolClass();
            simpleLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;
            simpleLineSymbol.Color = rgbColor;
            simpleLineSymbol.Width = 4;
            IWorkspace pLWorkspace;
            pLWorkspace = FileGDBWorkspaceFromPath(GIS_Settings.GISRootFolder);
            IFeatureWorkspace pFeatureWorkspace = pLWorkspace as IFeatureWorkspace;
            IFeatureClass pLineFeatureClass = pFeatureWorkspace.OpenFeatureClass(TLayer);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            int nameRes = pLineFeatureClass.FindField(_layers.AttributeIndexName(TLayer));

            // Clear the Selected Items List (if needed) and add the newly selected items
            _SelectedFeatures.Clear();

            using (ComReleaser comReleaser = new ComReleaser())
            {
                //IFeature feature = null;
                int index = GetIndexNumberFromLayerName(TLayer);
                const int maxConditionsPerStatement = 10;
                IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
                if (ClearExisting)
                    ((IFeatureSelection)featureLayer).Clear();
                IFeatureSelection featureSelection = featureLayer as IFeatureSelection;
                int counter = 0;
                int batchcounter = 0;
                string logicaloperator = "";
                string whereclause = "(";
                string dbwhereclause = "(";
                string attindexname = _layers.AttributeIndexName(TLayer);
                string dbindexname = _layers.CWIndexName(TLayer);

                // Build the Where Clause to search with using all the keys and OR'ing them

                while (counter <= (KeyIDs.Count - 1))
                {
                    if ((batchcounter > 0) && (counter < (KeyIDs.Count - 1)))
                        logicaloperator = " OR "; // Only use when not the first or last condition
                    if (KeyIDs[counter] != null)
                    {
                        whereclause += string.Format("{0} ({1} = {2})", logicaloperator, attindexname, KeyIDs[counter]);
                        dbwhereclause += string.Format("{0} ([{1}] = {2})", logicaloperator, dbindexname, KeyIDs[counter]);
                    }
                    if ((batchcounter > maxConditionsPerStatement) || (counter == (KeyIDs.Count - 1))) // if maxstatements OR last item
                    {
                        whereclause += ")";
                        dbwhereclause += ")";
                        // Set the filter
                        queryFilter.WhereClause = whereclause;
                        IFeatureCursor cursor = pLineFeatureClass.Search(queryFilter, true);
                        comReleaser.ManageLifetime(cursor);
                        IFeature pFeature = cursor.NextFeature();
                        while (pFeature != null)
                        {
                            featureSelection.Add(pFeature);
                            pEnvelope.Union(pFeature.Shape.Envelope);
                            retcount++;
                            pFeature = cursor.NextFeature();
                        }

                        LayerDef ld = new LayerDef();
                        _layers.GetLayerDef(TLayer, ref ld);
                        _cwdata.LoadSelectedItemsList(dbwhereclause, ld, false, ref _SelectedFeatures);

                        // Clear the values for the next batch
                        logicaloperator = "";
                        whereclause = "(";
                        dbwhereclause = "(";
                        batchcounter = 0;
                    }
                    else
                        batchcounter++;
                    counter++;
                } // while
                featureSelection.SelectionSymbol = simpleLineSymbol as ISymbol;
                _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
            }
            return retcount;
        }



        //-----------------------------------------------------------------------------------------------
        // Select the Features from the specified Layer that do not have a CW Index field value or it is = 0
        //-----------------------------------------------------------------------------------------------
        public int select_SelectUnlinkedFeaturesInMap(ref List<SelectedItemData> SelFeatures, ref EnvelopeClass pEnvelope, string TLayer)
        {
            int retcount = 0;
            IColor rgbColor = new RgbColorClass { Red = 100 };
            ESRI.ArcGIS.Display.ISimpleLineSymbol simpleLineSymbol = new ESRI.ArcGIS.Display.SimpleLineSymbolClass();
            simpleLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;
            simpleLineSymbol.Color = rgbColor;
            simpleLineSymbol.Width = 4;
            IWorkspace pLWorkspace;
            pLWorkspace = FileGDBWorkspaceFromPath(GIS_Settings.GISRootFolder);
            IFeatureWorkspace pFeatureWorkspace = pLWorkspace as IFeatureWorkspace;
            IFeatureClass pLineFeatureClass = pFeatureWorkspace.OpenFeatureClass(TLayer);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            int nameRes = pLineFeatureClass.FindField(_layers.AttributeIndexName(TLayer));

            // Clear the Selected Items List (if needed) and add the newly selected items

            using (ComReleaser comReleaser = new ComReleaser())
            {
                //IFeature feature = null;
                int index = GetIndexNumberFromLayerName(TLayer);
                const int maxConditionsPerStatement = 10;
                IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
                //((IFeatureSelection)featureLayer).Clear(); 

                IFeatureSelection featureSelection = featureLayer as IFeatureSelection;
                int counter = 0;
                int batchcounter = 0;
                string logicaloperator = "";
                string whereclause = "(";
                string attindexname = _layers.AttributeObjectIndexName(TLayer);
                string dbindexname = _layers.CWIndexName(TLayer);

                // Build the Where Clause to search with using all the keys and OR'ing them

                while (counter <= (SelFeatures.Count - 1))
                {
                    if ((batchcounter > 0) && (counter <= (SelFeatures.Count - 1)))
                        logicaloperator = " OR "; // Only use when not the first or last condition 
                    whereclause += string.Format("{0} ({1} = {2})", logicaloperator, attindexname, SelFeatures[counter].AttribObjectID);
                    if ((batchcounter > maxConditionsPerStatement) || (counter == (SelFeatures.Count - 1))) // if maxstatements OR last item
                    {
                        whereclause += ")";
                        // Set the filter
                        queryFilter.WhereClause = whereclause;
                        IFeatureCursor cursor = pLineFeatureClass.Search(queryFilter, true);
                        comReleaser.ManageLifetime(cursor);
                        IFeature pFeature = cursor.NextFeature();
                        while (pFeature != null)
                        {
                            featureSelection.Add(pFeature);
                            pEnvelope.Union(pFeature.Shape.Envelope);
                            retcount++;
                            pFeature = cursor.NextFeature();
                        }
                        // Clear the values for the next batch
                        logicaloperator = "";
                        whereclause = "(";
                        batchcounter = 0;
                    }
                    else
                        batchcounter++;
                    counter++;
                } // while

                featureSelection.SelectionSymbol = simpleLineSymbol as ISymbol;
                _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
            }
            return retcount;
        }


        //-----------------------------------------------------------------------------------------------
        // Perform Lines in Systems and display
        //-----------------------------------------------------------------------------------------------
        public void line_SelectLinesInMap(List<string> LineIDs, ref EnvelopeClass pEnvelope, string TLayer)
        {
            IColor rgbColor = new RgbColorClass { Red = 100 };
            ESRI.ArcGIS.Display.ISimpleLineSymbol simpleLineSymbol = new ESRI.ArcGIS.Display.SimpleLineSymbolClass();
            simpleLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid;
            simpleLineSymbol.Color = rgbColor;
            simpleLineSymbol.Width = 4;
            IWorkspace pLWorkspace;
            pLWorkspace = FileGDBWorkspaceFromPath(GIS_Settings.GISRootFolder);
            IFeatureWorkspace pFeatureWorkspace = pLWorkspace as IFeatureWorkspace;
            IFeatureClass pLineFeatureClass = pFeatureWorkspace.OpenFeatureClass(TLayer);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            int nameRes = pLineFeatureClass.FindField(_layers.AttributeIndexName(TLayer));
            using (ComReleaser comReleaser = new ComReleaser())
            {
                int index = GetIndexNumberFromLayerName(TLayer);
                IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
                IFeatureSelection featureSelection = featureLayer as IFeatureSelection;
                int counter = 0;
                //int MaxL = LineIDs.Count;
                for (counter = 0; counter <= (LineIDs.Count - 1); counter++)
                {
                    if (LineIDs[counter] != null)
                    {
                    // Set the filter to return only the correct Line
                    queryFilter.WhereClause = _layers.AttributeIndexName(TLayer) + " = " + LineIDs[counter];
                    IFeatureCursor cursor = pLineFeatureClass.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IFeature pFeature =cursor.NextFeature();
                    if (pFeature != null) 
                    {         
                    featureSelection.Add(pFeature);
                    pEnvelope.Union(pFeature.Shape.Envelope);
                    }
                    }
                    featureSelection.SelectionSymbol = simpleLineSymbol as ISymbol;
                    _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
                }   
            }   
        }


        // ------------------------------------------------
        // Get the index number for the specified layer name.</summary>
        // ------------------------------------------------
        public bool SetSelectionSetFeatureDescriptions()
        {
            if (_cwdata != null)
                return _cwdata.SetSelectionSetFeatureDescriptions(_SelectedFeatures, _layers);
            else
                return false;
        }

        // ------------------------------------------------
        // Get the index number for the specified layer name.</summary>
        // ------------------------------------------------
        public System.Int32 GetIndexNumberFromLayerName(System.String layerName)
        {
            if (_activeview== null || layerName == null)
            {
                return -1;
            }
            ESRI.ArcGIS.Carto.IMap map = _activeview.FocusMap;

            // Get the number of layers
            int numberOfLayers = map.LayerCount;

            // Loop through the layers and get the correct layer index
            for (System.Int32 i = 0; i < numberOfLayers; i++)
            {
                if (layerName == map.get_Layer(i).Name)
                {

                    // Layer was found
                    return i;
                }
            }// No layer was found
            return -1;
        }



        // ------------------------------------------------
        // Set Balloon for a point
        // ------------------------------------------------
        public void SetBalloonPoint(string locationid, string layername)
        {
            //IMxDocument pmxdocument = (IMxDocument)ArcMap.Application.Document;
            ITextElement ptextelement = new TextElementClass();
            IElement pelement = (IElement)ptextelement;
            ptextelement.Text = locationid;
            ESRI.ArcGIS.Carto.IMap map = _activeview.FocusMap;
            // Clear any previous buffers from the screen
            ESRI.ArcGIS.Carto.IGraphicsContainer graphicsContainer = (ESRI.ArcGIS.Carto.IGraphicsContainer)map; // Explicit Cast
            graphicsContainer.DeleteAllElements();

            IPoint LocXY = null;
            GetFeatureLocationXYPoint(locationid, layername, ref LocXY);
            double mdx = LocXY.X;//-88.667758;//LocXY.X;// ((pActiveView.Extent).XMax + (pActiveView.Extent.XMin)) / 2;
            double mdy = LocXY.Y;//41.246114; //LocXY.Y ;// ((pActiveView.Extent).YMax + (pActiveView.Extent.YMin)) / 2;

            IPoint point = new Point();
            point.PutCoords((mdx - (_activeview.Extent.Height) / 6), (mdy + (_activeview.Extent.Width) / 8));
            //point.PutCoords((mdx -1700/200), (mdy +1700/400));
            //point.PutCoords((mdx - 10), (mdy + 10));
            pelement.Geometry = point;

            IFormattedTextSymbol ptexsymbol = new TextSymbolClass();
            ICallout pcallout = new BalloonCallout();
            ptexsymbol.Background = (ITextBackground)pcallout;

            point.PutCoords((mdx), (mdy));
            pcallout.AnchorPoint = point;
            ptextelement.Symbol = ptexsymbol;

            IGraphicsContainer pGraphicsContainer = (IGraphicsContainer)_activeview;
            pGraphicsContainer.AddElement(pelement, 0);
            pelement.Activate(_activeview.ScreenDisplay);
            _activeview.PartialRefresh(esriViewDrawPhase.esriViewGraphics, pelement, null);
        }


        // ------------------------------------------------
        // Set Balloon for a Line or curve
        // ------------------------------------------------
        public void SetBalloonPolyLine(string locationid, string layername, string featureName)
        {
            //IMxDocument pmxdocument = (IMxDocument)ArcMap.Application.Document;
            ITextElement ptextelement = new TextElementClass();
            IElement pelement = (IElement)ptextelement;
            ptextelement.Text = featureName;
            ESRI.ArcGIS.Carto.IMap map = _activeview.FocusMap;
            // Clear any previous buffers from the screen
            ESRI.ArcGIS.Carto.IGraphicsContainer graphicsContainer = (ESRI.ArcGIS.Carto.IGraphicsContainer)map; // Explicit Cast
            graphicsContainer.DeleteAllElements();

            IPoint LocXY = null;

            if (_layers.KeyFieldIsString(layername))
            {
                GetFeatureLocationXYPolyLine(locationid, layername, ref LocXY);
            }
            else
            {
                GetFeatureLocationXYPolyLine(Convert.ToInt32(locationid), layername, ref LocXY);
            }

            double mdx = LocXY.X;
            double mdy = LocXY.Y;

            IPoint point = new Point();
            point.PutCoords((mdx - (_activeview.Extent.Height) / 6), (mdy + (_activeview.Extent.Width) / 8));
            pelement.Geometry = point;


            IFormattedTextSymbol ptexsymbol = new TextSymbolClass();
            ICallout pcallout = new BalloonCallout();
            ptexsymbol.Background = (ITextBackground)pcallout;

            point.PutCoords((mdx), (mdy));
            pcallout.AnchorPoint = point;
            ptextelement.Symbol = ptexsymbol;

            IGraphicsContainer pGraphicsContainer = (IGraphicsContainer)_activeview;
            pGraphicsContainer.AddElement(pelement, 0);
            pelement.Activate(_activeview.ScreenDisplay);
            _activeview.PartialRefresh(esriViewDrawPhase.esriViewGraphics, pelement, null);
        }



        // ------------------------------------------------
        // Point Coordinates
        // ------------------------------------------------
        private void GetFeatureLocationXYPoint(string objectID, string layername, ref IPoint LocXY)
        {
            IQueryFilter queryfilter = new QueryFilter();

            queryfilter.WhereClause = _layers.AttributeIndexName(layername) +  " = '" + objectID + "'";

            int index = GetIndexNumberFromLayerName(layername);
            IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
            IFeatureSelection featureSelection = featureLayer as IFeatureSelection;
            featureSelection.SelectFeatures(queryfilter, ESRI.ArcGIS.Carto.esriSelectionResultEnum.esriSelectionResultNew, false);
            IFeatureCursor pcursor = featureLayer.Search(queryfilter, false);
            IFeature pFeature = pcursor.NextFeature();
            IPoint point = null;
            point = (IPoint)pFeature.Shape;
            LocXY = point;
            _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
        }



        // ------------------------------------------------
        // Point Coordinates
        // Overload 1
        // String ObjectID
        // ------------------------------------------------
        private void GetFeatureLocationXYPolyLine(string objectID, string layername, ref IPoint LocXY)
        {
            IQueryFilter queryfilter = new QueryFilter();

            queryfilter.WhereClause = _layers.AttributeIndexName(layername) + " = '" + objectID + "'";
            int index = GetIndexNumberFromLayerName(layername);
            IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
            IFeatureSelection featureSelection = featureLayer as IFeatureSelection;
            featureSelection.SelectFeatures(queryfilter, ESRI.ArcGIS.Carto.esriSelectionResultEnum.esriSelectionResultNew, false);
            IFeatureCursor pcursor = featureLayer.Search(queryfilter, false);
            IFeature pFeature = pcursor.NextFeature();
            //point = (IPoint)pFeature.Shape;
            ICurve curve = pFeature.Shape as ICurve;
            IPoint curveMidPoint = new ESRI.ArcGIS.Geometry.Point();
            curve.QueryPoint(esriSegmentExtension.esriNoExtension, 0.5, true, curveMidPoint);
            //LocXY = point;
            LocXY = curveMidPoint;
            _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
        }

        // ------------------------------------------------
        // Point Coordinates 
        // Overload 2
        // integer ObjectID
        // ------------------------------------------------
        private void GetFeatureLocationXYPolyLine(int objectID, string layername, ref IPoint LocXY)
        {

            IQueryFilter queryfilter = new QueryFilter();

            queryfilter.WhereClause = _layers.AttributeIndexName(layername) + " = " + objectID;

            int index = GetIndexNumberFromLayerName(layername);
            IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
            IFeatureSelection featureSelection = featureLayer as IFeatureSelection;
            featureSelection.SelectFeatures(queryfilter, ESRI.ArcGIS.Carto.esriSelectionResultEnum.esriSelectionResultNew, false);
            IFeatureCursor pcursor = featureLayer.Search(queryfilter, false);
            IFeature pFeature = pcursor.NextFeature();
            //point = (IPoint)pFeature.Shape;
            ICurve curve = pFeature.Shape as ICurve;
            IPoint curveMidPoint = new ESRI.ArcGIS.Geometry.Point();
            curve.QueryPoint(esriSegmentExtension.esriNoExtension, 0.5, true, curveMidPoint);
            //LocXY = point;
            LocXY = curveMidPoint;
            _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
        }


        //-----------------------------------------------------------------------------------------------
        // Draws graphic buffers around the selected features in the map using distance units specified.</summary>
        //-----------------------------------------------------------------------------------------------
        public void CreateGraphicBuffersAroundSelectedFeatures(System.Double distance)
        {
            //parameter check
            if (_activeview == null || distance < 0)
            {
                return;
            }
            IMap map = _activeview.FocusMap;
            //// Clear any previous buffers from the screen
            IGraphicsContainer graphicsContainer = (ESRI.ArcGIS.Carto.IGraphicsContainer)map; // Explicit Cast
            //graphicsContainer.DeleteAllElements();

            // Verify there is a feature(s) selected
            if (map.SelectionCount == 0)
            {
                return;
            }
            // Reset to the first selected feature
            IEnumFeature enumFeature = (IEnumFeature)map.FeatureSelection; // Explicit Cast
            enumFeature.Reset();
            IFeature feature = enumFeature.Next();
            // Buffer all the selected features by the buffer distance and create a new polygon element from each result
            ITopologicalOperator topologicalOperator;
            IElement element;
            while (!(feature == null))
            {
                topologicalOperator = (ITopologicalOperator)feature.Shape; // Explicit Cast
                element = new PolygonElementClass();
                element.Geometry = topologicalOperator.Buffer(distance);
                ISimpleFillSymbol simpleFill = new SimpleFillSymbolClass();
                simpleFill.Color = new RgbColorClass() { IColor_NullColor = true };
                simpleFill.Outline = new SimpleLineSymbolClass() { Color = new RgbColorClass() { Red = 255 } };
                ((IFillShapeElement)element).Symbol = simpleFill;
                graphicsContainer.AddElement(element, 0);
                feature = enumFeature.Next();
            }
            _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGraphics, null, null);
        }


        //-----------------------------------------------------------------------------------------------
        // Returns CompWorksID from Feature ObjectID
        // Overload 1 
        //-----------------------------------------------------------------------------------------------
        public string GetCWID_From_AttributeID(int objectid, string layername)
        {
            string id = "";
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.SubFields = _layers.AttributeIndexName(layername) + ", " + pTableOwner.OIDFieldName;
            queryFilter.WhereClause = pTableOwner.OIDFieldName + " = " + objectid;
            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            int ObjectIndex = pTableOwner.FindField(_layers.AttributeIndexName(layername));
            using (ComReleaser comReleaser = new ComReleaser())
                try
                {
                    ICursor cursor = pTableOwner.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IRow row = null;
                    while ((row = cursor.NextRow()) != null)
                    {
                        id = row.get_Value(ObjectIndex).ToString();
                        //String Location = Convert.ToString(row.get_Value(LocationIndex));
                    }
                }
                catch (Exception err)
                {
                    string M = err.Message;
                }
            // pass it back
            return id;
        }


        //-----------------------------------------------------------------------------------------------
        // Returns true if the specified field exists in the Layer's Attribute Table Fields
        //-----------------------------------------------------------------------------------------------
        public bool LayerAttributeFieldExists(string layername, string fieldName)
        {
            List<string> FieldList = new List<string>();
            if (GetLayerAttributeFields(layername, ref FieldList) > 0)
            {
                return (FieldList.IndexOf(fieldName) >= 0) ;
            }
            return false;
        }



        //-----------------------------------------------------------------------------------------------
        // Returns the Layer's Attribute Table Field Names
        //-----------------------------------------------------------------------------------------------
        public int GetLayerAttributeFields(string layername, ref List<string> FieldList)
        {
            int retval = -1;
            if ((FieldList == null) || (layername.Length <= 0))
                return retval;
            FieldList.Clear();
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            ITable pTableOwner;
            try
            {
                pTableOwner = pFeatureWorkspace.OpenTable(layername);
                int i;
                for (i = 0; i <= (pTableOwner.Fields.FieldCount - 1); i++)
                {
                    FieldList.Add(pTableOwner.Fields.Field[i].Name);
                }
                retval = FieldList.Count;
            }
            catch (Exception err)
            {
                if (err.Message.Length > 0)
                {
                    retval = -2;
                }
            }
            return retval;
        }




        //-----------------------------------------------------------------------------------------------
        // Returns CompWorksID from Feature ObjectID
        // Overload 2  - returns Item Description
        //-----------------------------------------------------------------------------------------------
        public string GetCWID_From_AttributeID(int objectid, string layername, ref string itemdesc)
        {
            string id = "";
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.SubFields =  string.Format("{0}, {1}, {2}",_layers.AttributeIndexName(layername),pTableOwner.OIDFieldName,_layers.CWFeatureDisplayField(layername));
            queryFilter.WhereClause = pTableOwner.OIDFieldName + " = " + objectid;
            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            int ObjectIndex = pTableOwner.FindField(_layers.AttributeIndexName(layername));
            using (ComReleaser comReleaser = new ComReleaser())
                try
                {
                    ICursor cursor = pTableOwner.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IRow row = null;
                    while ((row = cursor.NextRow()) != null)
                    {
                        id = row.get_Value(ObjectIndex).ToString();
                        //String Location = Convert.ToString(row.get_Value(LocationIndex));
                    }
                }
                catch (Exception err)
                {
                    string M = err.Message;
                }
            // pass it back
            return id;
        }


        //-----------------------------------------------------------------------------------------------
        // Returns CompWorksID from Feature ObjectID
        //-----------------------------------------------------------------------------------------------
        public string GetCWID_From_AttributeID_Str(int objectid, string layername)
        {
            string id = "";
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.SubFields = _layers.AttributeIndexName(layername) + ", " + pTableOwner.OIDFieldName;
            queryFilter.WhereClause = pTableOwner.OIDFieldName + " = " + objectid;
            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            int ObjectIndex = pTableOwner.FindField(_layers.AttributeIndexName(layername));
            using (ComReleaser comReleaser = new ComReleaser())
                try
                {
                    ICursor cursor = pTableOwner.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IRow row = null;
                    while ((row = cursor.NextRow()) != null)
                    {
                        id = row.get_Value(ObjectIndex).ToString();
                    }
                }
                catch (Exception err)
                {
                    string M = err.Message;
                }
            // pass it back
            return id;
        }



        ///<summary>Select features in the IActiveView by an attribute query using a SQL syntax in a where clause.</summary>
        /// 
        ///<param name="activeView">An IActiveView interface</param>
        ///<param name="featureLayer">An IFeatureLayer interface to select upon</param>
        ///<param name="whereClause">A System.String that is the SQL where clause syntax to select features. Example: "CityName = 'Redlands'"</param>
        ///  
        ///<remarks>Providing and empty string "" will return all records.</remarks>
        public void SelectMapFeaturesByAttributeQuery(ESRI.ArcGIS.Carto.IActiveView activeView, ESRI.ArcGIS.Carto.IFeatureLayer featureLayer, System.String whereClause)
        {
            if (activeView == null || featureLayer == null || whereClause == null)
            {
                return;
            }
            ESRI.ArcGIS.Carto.IFeatureSelection featureSelection = featureLayer as ESRI.ArcGIS.Carto.IFeatureSelection; // Dynamic Cast

            // Set up the query
            ESRI.ArcGIS.Geodatabase.IQueryFilter queryFilter = new ESRI.ArcGIS.Geodatabase.QueryFilterClass();
            queryFilter.WhereClause = whereClause;

            // Invalidate only the selection cache. Flag the original selection
            activeView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, null, null);

            // Perform the selection
            featureSelection.SelectFeatures(queryFilter, ESRI.ArcGIS.Carto.esriSelectionResultEnum.esriSelectionResultNew, false);

            // Flag the new selection
            activeView.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewGeoSelection, null, null);
        }


        //-----------------------------------------------------------------------------------------------
        // Workspace
        //-----------------------------------------------------------------------------------------------
        protected static IWorkspace FileGDBWorkspaceFromPath(String path)
        {
            Type t = Type.GetTypeFromProgID("esriDataSourcesGDB.FileGDBWorkspaceFactory");
            //System.Object obj = Activator.CreateInstance(t);

            IWorkspaceFactory workspaceFactory = Activator.CreateInstance(t) as IWorkspaceFactory;
            return workspaceFactory.OpenFromFile(path, 0);
        }



        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }



        //-----------------------------------------------------------------------------------------------
        // Sends the Selected Items List to Excel
        //-----------------------------------------------------------------------------------------------
        public bool SendSelectedItemsToExcel(string ExcelFileName)
        {

            if (_SelectedFeatures.Count <= 0) return false;
            if (ExcelFileName.Length == 0) return false;
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            Int16 i, j;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            // Set Title Rows
            xlWorkSheet.Cells[1, 1] = "Layer";
            xlWorkSheet.Cells[1, 2] = "Feature Name";
            // Add Features
            for (i = 0; i <= _SelectedFeatures.Count - 1; i++)
            {
                // Now Write Rows/Columns
                xlWorkSheet.Cells[i + 2, 1] = _SelectedFeatures[i].layername;
                xlWorkSheet.Cells[i + 2, 2] = _SelectedFeatures[i].featuredescription;
            }

            xlApp.DisplayAlerts = false;
            xlWorkBook.SaveAs(@ExcelFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            // Release the objects 
            // Then do Garbage Collection
            try
            {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            xlWorkSheet = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            xlWorkBook = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            xlApp = null;
            }
            catch (Exception ex)
            {
               // MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
            return true;
        }


        //-----------------------------------------------------------------------------------------------
        // loads the GIS data
        // pass the mxd to open
        //-----------------------------------------------------------------------------------------------
        public bool Loadgis(string gisdbpath)
        {
            return true;
        }


        //-----------------------------------------------------------------------------------------------
        // Loads the GIS configuration settings from the CompWorks BP database
        //-----------------------------------------------------------------------------------------------
        public bool LoadConfiguration(string cwdatafile)
        {
            // Load the configuration
            return true;
        }


        //-----------------------------------------------------------------------------------------------
        // Query the WorkSpace and obtain the ObjectID Field names for the System Layers
        //-----------------------------------------------------------------------------------------------
        public bool LoadLayerObjectIDFields(IWorkspace pWorkspace, LayerDefs layers)
        {
            int i;          
            IFeatureWorkspace pFeatureWorkspace = pWorkspace as IFeatureWorkspace;
            LayerDef ldef = new LayerDef();
            try
            {
                for (i = 0; i <= (layers.Count - 1); i++)
                {
                    if (_layers.GetLayerDefOrdinal(i, ref ldef) == true)
                        if (ldef.SystemUse > 0)
                        {
                            ldef.LayerTableName = getLayerTableName(ldef.layername);
                            ITable pTableOwner = pFeatureWorkspace.OpenTable(ldef.LayerTableName);
                            ldef.GIS_Attribute_Table_Object_Index_Field = pTableOwner.OIDFieldName;
                        }
                }
            }
            catch
            {
                return false;
            }
            return true;
        }


        //-----------------------------------------------------------------------------------------------
        // Sets the Layer's Renderer
        //-----------------------------------------------------------------------------------------------
        public bool SetLayerRenderer(LayerDef ldef)
        {
            //Get the layer selected in the combo box
            IGeoFeatureLayer geofeaturelayer = null;
            IMap map = _activeview.FocusMap;
            for (int i = 0; i <= map.LayerCount - 1; i++)
            {
                if (map.get_Layer(i).Name == ldef.layername)
                {
                    geofeaturelayer = (IGeoFeatureLayer)map.get_Layer(i);
                    break;
                }
            }
            if (geofeaturelayer == null) return false;
            //Set the new renderer 
            if (ldef.m_classBreaksRenderer != null)
            {
                geofeaturelayer.Renderer = (IFeatureRenderer)ldef.m_classBreaksRenderer;
                //Trigger contents changed event for TOCControl
                _activeview.ContentsChanged();
                //Refresh the display 
                _activeview.PartialRefresh(esriViewDrawPhase.esriViewGeography, geofeaturelayer, null);
            }
            return true;
        }


        public static IRgbColor MakeRGBColor(byte R, byte G, byte B)
        {
            ESRI.ArcGIS.Display.RgbColor RgbClr = new ESRI.ArcGIS.Display.RgbColorClass();
            RgbClr.Red = R;
            RgbClr.Green = G;
            RgbClr.Blue = B;
            return RgbClr;
        }


        //-----------------------------------------------------------------------------------------------
        // Sets the Layer's Renderer to Basic Line
        //-----------------------------------------------------------------------------------------------
        public bool SetLayerRendererBasicLine(LayerDef ldef)
        {
            if (ldef.m_SimpleRenderer == null) return false;
            IGeoFeatureLayer geofeaturelayer = null;
            IMap map = _activeview.FocusMap;
            for (int i = 0; i <= map.LayerCount - 1; i++)
            {
                if (map.get_Layer(i).Name == ldef.layername)
                {
                    geofeaturelayer = (IGeoFeatureLayer)map.get_Layer(i);
                    break;
                }
            }
            if (geofeaturelayer == null) return false;
            //Set the new renderer 
            if (ldef.m_SimpleRenderer != null)
            {
                geofeaturelayer.Renderer = (IFeatureRenderer)ldef.m_SimpleRenderer;
                //Trigger contents changed event for TOCControl
                _activeview.ContentsChanged();
                //Refresh the display 
                _activeview.PartialRefresh(esriViewDrawPhase.esriViewGeography, geofeaturelayer, null);
            }
            return true;
        }


        //-----------------------------------------------------------------------------------------------
        // Check Database for Field Updates
        //-----------------------------------------------------------------------------------------------
        //public bool DoFieldUpdates()
        //{
        //    string indexFieldName = GetCWAttributeIndexFieldName("");
        //    return true;
        //}



        //-----------------------------------------------------------------------------------------------
        // return the name of a layer's table (most are the same as the layer name)
        //-----------------------------------------------------------------------------------------------
        public string getLayerTableName(string layerName)
        {
            string tablename = "";
            if ((_activeview == null) || (layerName.Length == 0)) return tablename;
            int index = GetIndexNumberFromLayerName(layerName);
            IFeatureLayer fl = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
            tablename = ((IDataset)fl.FeatureClass).BrowseName;
            return tablename;
        }



        //-----------------------------------------------------------------------------------------------
        // class constructor
        // pass the mx datafile to open
        //-----------------------------------------------------------------------------------------------
        public ArcEngineClass(string cwdatafile, string cwsystemfile)
        {
            _initializeresult = 0;
            _initializeerror = "";
            // Connect to the CW Database
            _cwdata = new CompWorksDBConnect(cwdatafile, cwsystemfile);
            // Get the GIS settings from the conifiguration
            if (_cwdata.setup_Get_GIS_Settings(ref GIS_Settings) != true)
            {
                _initializeresult = 101;
                _initializeerror = "Unable to load the GIS Settings in CompWorks!";
                return;
            }

            // Load the CompWorks definition records for the Layers
            _layers = new LayerDefs();
            if (_cwdata.SetLayerDefs(_layers) != true)
            {
                _initializeresult = 102;
                _initializeerror = "Unable to load the Layer Definitions in CompWorks!";
                return;
            }
            _cwdatafile = cwdatafile;
            _cwsystemfile = cwsystemfile;
            _workspace = FileGDBWorkspaceFromPath(GIS_Settings.GISRootFolder);

            // Identify and set the Attribue Index field for the System Layers
            //if (LoadLayerObjectIDFields(_workspace, _layers) != true)
            //{
            //    _initializeresult = 103;
            //    _initializeerror = "Unable to load the Layer Object IDs!";
            //    return;
            //}
            _SelectedFeatures = new List<SelectedItemData>();
        } // Constructor

    }
}
