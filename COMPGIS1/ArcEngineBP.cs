using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.ADF;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.esriSystem;

namespace COMPGIS1
{
    public class ArcEngineBP : ArcEngineClass 
    {


        //-----------------------------------------------------------------------------------------------
        // class constructor
        // pass the mx datafile to open
        //-----------------------------------------------------------------------------------------------
        public ArcEngineBP(string cwdatafile, string cwsystemfile) : base(cwdatafile, cwsystemfile)
        {
          //
        }


        //-----------------------------------------------------------------------------------------------
        // Run Find - Inspected Line Sections
        // Parameter is the SetupID from the Dbf Setup Data record for the "MethodType"
        //-----------------------------------------------------------------------------------------------
        public bool find_ShowInspectedLineSections(int InspectionSetupID)
        {
            if (InspectionSetupID <= 0) return false;
            // Get the Layer Def
            LayerDef ldef = new LayerDef();
            _layers.GetLayerDef(LayerSystemUse.suBPLineSections, ref ldef);
            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            _SelectedFeatures.Clear();
            if (_cwdata.find_Line_Section_Inspections(ref _SelectedFeatures, ldef, InspectionSetupID) == true)
            {
                select_ClearAllSelections();
                if (select_SelectFeaturesInMap(ref _SelectedFeatures, ref pEnvelope, ldef.layername) > 0)
                {
                    SetZoomEnvelope(ref pEnvelope);
                    //                    pEnvelope.Expand(1.2, 1.2, true);
                    _activeview.Extent = pEnvelope;
                    _activeview.Refresh();
                    _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
                    return true;
                }
            }
            return false;
        }


        //-----------------------------------------------------------------------------------------------
        // Run Find - Evaluated Line Sections
        // Parameter is the SetupID from the Dbf Setup Data record for the "MethodType"
        //-----------------------------------------------------------------------------------------------
        public bool find_ShowEvaluatedLineSections(string Method, string WorkScope, bool applytocurrent = false)
        {
            // Get the Layer Def
            LayerDef ldef = new LayerDef();
            _layers.GetLayerDef(LayerSystemUse.suBPLineSections, ref ldef);
            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            List<SelectedItemData> tempList = new List<SelectedItemData>();
            List<SelectedItemData> resultList = new List<SelectedItemData>();
            if (_cwdata.find_Line_Section_Evaluations(ref tempList, ldef, Method, WorkScope) != true)
            {
                return false;
            }
            if (applytocurrent)
            {
                joinlists_Inner(tempList, _SelectedFeatures, ref resultList);
                lists_Assign(resultList, ref _SelectedFeatures);
                resultList.Clear();
                tempList.Clear();
            }
            else
            {
                lists_Assign(tempList, ref _SelectedFeatures);
                tempList.Clear();
            }
            if (!applytocurrent)
                select_ClearAllSelections();
            if (select_SelectFeaturesInMap(ref _SelectedFeatures, ref pEnvelope, ldef.layername) > 0)
            {
                SetZoomEnvelope(ref pEnvelope);
                _activeview.Extent = pEnvelope;
                _activeview.Refresh();
                _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
                return true;
            }
            return false;
        }


        //-----------------------------------------------------------------------------------------------
        // Run Find - Line Sections Associated with LCOs
        //-----------------------------------------------------------------------------------------------
        public bool find_ShowLineSectionsWithLCO(int LCOValue, bool inclusivebelow = false, bool applytocurrent = false)
        {
            // Get the Layer Def
            LayerDef ldef = new LayerDef();
            _layers.GetLayerDef(LayerSystemUse.suBPLineSections, ref ldef);
            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            if (!applytocurrent)
            {
                _SelectedFeatures.Clear();
            }
            List<SelectedItemData> tempList = new List<SelectedItemData>();
            List<SelectedItemData> resultList = new List<SelectedItemData>();
            if (_cwdata.find_Line_Sections_with_LCO(ref tempList, ldef, LCOValue, inclusivebelow) != true)
            {
                return false;
            }
            if (applytocurrent)
            {
                joinlists_Inner(tempList, _SelectedFeatures, ref resultList);
                lists_Assign(resultList, ref _SelectedFeatures);
                resultList.Clear();
                tempList.Clear();
            }
            else
            {
                lists_Assign(tempList, ref _SelectedFeatures);
                tempList.Clear();
            }
            if (!applytocurrent)
                select_ClearAllSelections();
            if (select_SelectFeaturesInMap(ref _SelectedFeatures, ref pEnvelope, ldef.layername) > 0)
            {
                SetZoomEnvelope(ref pEnvelope);
                _activeview.Extent = pEnvelope;
                _activeview.Refresh();
                _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
                return true;
            }
            return false;
        }

        //-----------------------------------------------------------------------------------------------
        // Run Find - Line Sections Associated with GWUT Inspections by Level
        //-----------------------------------------------------------------------------------------------
        public bool find_ShowLineSectionsWithGWUTByLevel(int Level, bool applytosel)
        {
            // Get the Layer Def
            LayerDef ldef = new LayerDef();
            _layers.GetLayerDef(LayerSystemUse.suBPLineSections, ref ldef);
            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            List<SelectedItemData> tempList = new List<SelectedItemData>();
            List<SelectedItemData> resultList = new List<SelectedItemData>();
            if (applytosel)
            {
                // Build temp list and merge results with existing
                if (_cwdata.find_Line_Section_GWUT_By_Level(ref tempList, ldef, Level) > 0)
                {
                    joinlists_Inner(tempList, _SelectedFeatures, ref resultList);
                    lists_Assign(resultList, ref _SelectedFeatures);
                    resultList.Clear();
                    tempList.Clear();
                }
                else
                {
                    return false;
                }
            }
            else
            {
                _SelectedFeatures.Clear();
                if (_cwdata.find_Line_Section_GWUT_By_Level(ref _SelectedFeatures, ldef, Level) <= 0)
                {
                    return false;
                }
            }

            if (!applytosel)
                select_ClearAllSelections();

            if (select_SelectFeaturesInMap(ref _SelectedFeatures, ref pEnvelope, ldef.layername) > 0)
            {
                SetZoomEnvelope(ref pEnvelope);
                _activeview.Extent = pEnvelope;
                _activeview.Refresh();
                _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
                return true;
            }
            return false;
        }


        //-----------------------------------------------------------------------------------------------
        // Run Find - Line Sections by Risk Ranking
        //-----------------------------------------------------------------------------------------------
        public bool find_ShowLineSectionsByRiskRanking(string sussql, string consql, bool applytosel)
        {
            // Get the Layer Def
            LayerDef ldef = new LayerDef();
            _layers.GetLayerDef(LayerSystemUse.suBPLineSections, ref ldef);
            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            List<SelectedItemData> _selfeaturestemp = new List<SelectedItemData>();
            List<SelectedItemData> _resfeaturestemp = new List<SelectedItemData>();
            _selfeaturestemp.Clear();
            if (_cwdata.find_Line_Section_by_Risk_Ranking(ref _selfeaturestemp, ldef, sussql, consql) > 0)
            {
                if (applytosel)
                {
                    joinlists_Inner(_selfeaturestemp, _SelectedFeatures, ref _resfeaturestemp);
                    lists_Assign(_resfeaturestemp, ref _SelectedFeatures);
                }
                else
                {
                    lists_Assign(_selfeaturestemp, ref _SelectedFeatures);
                }
                _resfeaturestemp.Clear();
                _selfeaturestemp.Clear();
                if (!applytosel)
                    select_ClearAllSelections();
                if (select_SelectFeaturesInMap(ref _SelectedFeatures, ref pEnvelope, ldef.layername) > 0)
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
        // Run Find - Lines with Licensed Material
        // Parameter is the SetupID from the Dbf Setup Data record for the "MethodType"
        //-----------------------------------------------------------------------------------------------
        public bool find_ShowLinesWithLicensedMaterial()
        {
            // Get the Layer Def
            LayerDef ldef = new LayerDef();
            _layers.GetLayerDef(LayerSystemUse.suBPLines, ref ldef);
            ESRI.ArcGIS.Geometry.EnvelopeClass pEnvelope;
            pEnvelope = new EnvelopeClass();
            _SelectedFeatures.Clear();
            if (_cwdata.find_Line_With_Licensed_Material(ref _SelectedFeatures, ldef) == true)
            {
                select_ClearAllSelections();
                if (select_SelectFeaturesInMap(ref _SelectedFeatures, ref pEnvelope, ldef.layername) > 0)
                {
                    SetZoomEnvelope(ref pEnvelope);
                    //                    pEnvelope.Expand(1.2, 1.2, true);
                    _activeview.Extent = pEnvelope;
                    _activeview.Refresh();
                    _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
                    return true;
                }
            }
            return false;
        }





        //-----------------------------------------------------------------------------------------------
        // return the ZoneIDs of the line sections intersecting within a specified radius of the BL
        // the list of ZoneIDs in a string separated by a TAB
        //-----------------------------------------------------------------------------------------------
        public int ls_dospatialintersection(string boringlocationid, double radius, ref EnvelopeClass pEnvelope, ref int[] SpatialArray)
        {
            IWorkspace pWorkspace;

            int ArrayCount = 0;

            pWorkspace = FileGDBWorkspaceFromPath(GIS_Settings.GISRootFolder);
            PerformRadiusIntersect(pWorkspace, radius, _layers.LayerName(LayerSystemUse.suBPBoringLocations), _layers.LayerName(LayerSystemUse.suBPLineSections), ref SpatialArray, ref ArrayCount, ref pEnvelope, boringlocationid);
            _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);

            // pass it back
            return ArrayCount;
        }


        //-----------------------------------------------------------------------------------------------
        // return the Keys of the items in the target layer (TLayer) intersecting within a specified distance of the SLayer Item (SKey)
        // as a list of Keys in a string separated by a TAB
        //-----------------------------------------------------------------------------------------------
        public int find_dospatialintersection(string SKey, string SLayer, string TLayer, double distance, ref EnvelopeClass pEnvelope, ref int[] SpatialArray)
        {
            IWorkspace pWorkspace;
            int ArrayCount = 0;
            pWorkspace = FileGDBWorkspaceFromPath(GIS_Settings.GISRootFolder);
            PerformRadiusIntersect(pWorkspace, distance, SLayer, TLayer, ref SpatialArray, ref ArrayCount, ref pEnvelope, SKey, false);
            _activeview.PartialRefresh(esriViewDrawPhase.esriViewAll, null, null);
            // pass it back
            return ArrayCount;
        }



        //-----------------------------------------------------------------------------------------------
        // Perform Lines select
        //-----------------------------------------------------------------------------------------------
        public void line_SelectLines(string[] LineIDs, ref EnvelopeClass pEnvelope, string TLayer, string featname)
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

            // Select the fields to be returned
            //int nameRes = pLineFeatureClass.FindField("LineID");
            using (ComReleaser comReleaser = new ComReleaser())
            {

                int index = GetIndexNumberFromLayerName(TLayer);
                IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
                string ssssss = featureLayer.Name;
                IFeatureSelection featureSelection = featureLayer as IFeatureSelection;

                int counter = 0;

                //int MaxL = LineIDs.Count;
                for (counter = 0; counter <= (LineIDs.Length - 1); counter++)
                {
                    if (LineIDs[counter] != null)
                    {
                        // Set the filter to return only the correct Line
                        queryFilter.WhereClause = _layers.AttributeIndexName(TLayer) + " = " + LineIDs[counter];
                        IFeatureCursor cursor = pLineFeatureClass.Search(queryFilter, true);
                        comReleaser.ManageLifetime(cursor);
                        IFeature pFeature = cursor.NextFeature();
                        if (pFeature != null)
                        {
                            featureSelection.Add(pFeature);
                            pEnvelope.Union(pFeature.Shape.Envelope);
                            //                            SetBalloonPolyLine(LineIDs[counter], TLayer, featname);

                        }
                    }
                    featureSelection.SelectionSymbol = simpleLineSymbol as ISymbol;
                    _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
                }
                //SetBalloon(ref pActiveView, objectID, layername);
            }
        }



        //-----------------------------------------------------------------------------------------------
        // Select Line Sections in Map
        //-----------------------------------------------------------------------------------------------
        public void line_SelectLineSections(string[] FeatureIDs, ref EnvelopeClass pEnvelope, string TLayer, string featname)
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

            // Select the fields to be returned
            //int nameRes = pLineFeatureClass.FindField("Zone_ID");
            using (ComReleaser comReleaser = new ComReleaser())
            {
                int index = GetIndexNumberFromLayerName(_layers.LayerName(LayerSystemUse.suBPLineSections));
                IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
                IFeatureSelection featureSelection = featureLayer as IFeatureSelection;
                int counter = 0;
                //int MaxL = LineIDs.Count;
                for (counter = 0; counter <= (FeatureIDs.Length - 1); counter++)
                {
                    if (FeatureIDs[counter] != null)
                    {
                        // Set the filter to return only the correct Line Section
                        queryFilter.WhereClause = _layers.AttributeIndexName(TLayer) + " = " + FeatureIDs[counter];
                        IFeatureCursor cursor = pLineFeatureClass.Search(queryFilter, true);
                        comReleaser.ManageLifetime(cursor);
                        IFeature pFeature = cursor.NextFeature();
                        if (pFeature != null)
                        {
                            featureSelection.Add(pFeature);
                            pEnvelope.Union(pFeature.Shape.Envelope);
                            //                            SetBalloonPolyLine(FeatureIDs[counter], TLayer, featname);
                        }
                    }
                    featureSelection.SelectionSymbol = simpleLineSymbol as ISymbol;
                    _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);

                }

                //// Set the filter to return only Line Sections.
                //featureSelection.SelectionColor = rgbColor;

            }
        }



        //-----------------------------------------------------------------------------------------------
        // Perform Boring Location select
        //-----------------------------------------------------------------------------------------------
        public void SelectBoringLocation(string[] FearureIDs, ref EnvelopeClass pEnvelope, string TLayer)
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

            // Select the fields to be returned
            //int nameRes = pLineFeatureClass.FindField("LineID");
            using (ComReleaser comReleaser = new ComReleaser())
            {

                int index = GetIndexNumberFromLayerName(TLayer);
                IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
                IFeatureSelection featureSelection = featureLayer as IFeatureSelection;

                int counter = 0;

                //int MaxL = FearureIDs.Count;
                for (counter = 0; counter <= (FearureIDs.Length - 1); counter++)
                {
                    if (FearureIDs[counter] != null)
                    {
                        // Set the filter to return only the correct Line

                        queryFilter.WhereClause = _layers.AttributeIndexName(TLayer) + " = '" + FearureIDs[counter] + "'";

                        IFeatureCursor cursor = pLineFeatureClass.Search(queryFilter, true);
                        comReleaser.ManageLifetime(cursor);
                        IFeature pFeature = cursor.NextFeature();
                        if (pFeature != null)
                        {
                            featureSelection.Add(pFeature);
                            pEnvelope.Union(pFeature.Shape.Envelope);
                            SetZoomEnvelope(ref pEnvelope);
                            //pEnvelope.XMax = pEnvelope.XMax + 5;
                            //pEnvelope.XMin = pEnvelope.XMin - 5;
                            //pEnvelope.YMax = pEnvelope.YMax + 5;
                            //pEnvelope.YMin = pEnvelope.YMin - 5;
                            //pEnvelope.Expand(1.2, 1.2, true);
                            //if (pEnvelope.XMax == pEnvelope.XMin) pEnvelope.XMax = pEnvelope.XMax + .00001;
                            //if (pEnvelope.YMax == pEnvelope.YMin) pEnvelope.YMax = pEnvelope.YMax + .00001;
                            _activeview.Extent = pEnvelope;
                            _activeview.Refresh();
                            _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
                            SetBalloonPoint(FearureIDs[counter], TLayer);
                        }
                    }
                    featureSelection.SelectionSymbol = simpleLineSymbol as ISymbol;
                    _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
                }
            }
        }


        //-----------------------------------------------------------------------------------------------
        // Perform CP Location select
        //-----------------------------------------------------------------------------------------------
        public void SelectCPLocation(string[] FearureIDs, ref EnvelopeClass pEnvelope, string TLayer)
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

            // Select the fields to be returned
            //int nameRes = pLineFeatureClass.FindField("LineID");
            using (ComReleaser comReleaser = new ComReleaser())
            {

                int index = GetIndexNumberFromLayerName(TLayer);
                IFeatureLayer featureLayer = _activeview.FocusMap.get_Layer(index) as IFeatureLayer;
                IFeatureSelection featureSelection = featureLayer as IFeatureSelection;

                int counter = 0;

                //int MaxL = FearureIDs.Count;
                for (counter = 0; counter <= (FearureIDs.Length - 1); counter++)
                {
                    if (FearureIDs[counter] != null)
                    {
                        // Set the filter to return only the correct Line


                        queryFilter.WhereClause = _layers.AttributeIndexName(TLayer) + " = '" + FearureIDs[counter] + "'";
                        IFeatureCursor cursor = pLineFeatureClass.Search(queryFilter, true);
                        comReleaser.ManageLifetime(cursor);
                        IFeature pFeature = cursor.NextFeature();
                        if (pFeature != null)
                        {
                            featureSelection.Add(pFeature);
                            pEnvelope.Union(pFeature.Shape.Envelope);
                            pEnvelope.XMax = pEnvelope.XMax + 5;
                            pEnvelope.XMin = pEnvelope.XMin - 5;
                            pEnvelope.YMax = pEnvelope.YMax + 5;
                            pEnvelope.YMin = pEnvelope.YMin - 5;
                            pEnvelope.Expand(1.2, 1.2, true);
                            if (pEnvelope.XMax == pEnvelope.XMin) pEnvelope.XMax = pEnvelope.XMax + .00001;
                            if (pEnvelope.YMax == pEnvelope.YMin) pEnvelope.YMax = pEnvelope.YMax + .00001;
                            _activeview.Extent = pEnvelope;
                            _activeview.Refresh();
                            _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);

                            SetBalloonPoint(FearureIDs[counter], TLayer);
                        }
                    }

                }
                featureSelection.SelectionSymbol = simpleLineSymbol as ISymbol;
                _activeview.PartialRefresh(ESRI.ArcGIS.Carto.esriViewDrawPhase.esriViewAll, null, null);
            }
        }


        //-----------------------------------------------------------------------------------------------
        // This function creates the cross reference between the ArcMAP feature ID (ObjectID) and
        // our CompWorks ID;
        // Dbf Lines table
        //-----------------------------------------------------------------------------------------------
        public bool linesUpdateCrossReference()
        {
            string layername = _layers.LayerName(LayerSystemUse.suBPLines);
            if (layername.Length == 0) return false;
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.SubFields = _layers.AttributeIndexName(layername) + ", " + pTableOwner.OIDFieldName;

            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            int CWFieldIndex = pTableOwner.FindField(_layers.AttributeIndexName(layername));
            string ffnn = pTableOwner.OIDFieldName;
            int FeatureFieldIndex = pTableOwner.FindField(ffnn);
            bool done = false;
            using (ComReleaser comReleaser = new ComReleaser())
                try
                {
                    ICursor cursor = pTableOwner.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IRow row = null;
                    while ((row = cursor.NextRow()) != null)
                    {
                        int cwid = Convert.ToInt32(row.get_Value(CWFieldIndex).ToString());
                        int fid = Convert.ToInt32(row.get_Value(FeatureFieldIndex).ToString());
                        done = _cwdata.lineSetRecordCrossReference(cwid, fid);
                    }
                }
                catch (Exception err)
                {
                    string M = err.Message;
                }
            return false;
        }



        //-----------------------------------------------------------------------------------------------
        // This function creates the cross reference between the ArcMAP feature ID (ObjectID) and
        // our CompWorks ID;
        // Dbf Line Sections table
        //-----------------------------------------------------------------------------------------------
        public bool linesectionUpdateCrossReference()
        {
            string layername = _layers.LayerName(LayerSystemUse.suBPLineSections);
            if (layername.Length == 0) return false;
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.SubFields = _layers.AttributeIndexName(layername) + ", " + pTableOwner.OIDFieldName;

            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            int CWFieldIndex = pTableOwner.FindField(_layers.AttributeIndexName(layername));
            int FeatureFieldIndex = pTableOwner.FindField(pTableOwner.OIDFieldName);
            bool done = false;
            using (ComReleaser comReleaser = new ComReleaser())
                try
                {
                    ICursor cursor = pTableOwner.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IRow row = null;
                    while ((row = cursor.NextRow()) != null)
                    {
                        int cwid = Convert.ToInt32(row.get_Value(CWFieldIndex).ToString());
                        int fid = Convert.ToInt32(row.get_Value(FeatureFieldIndex).ToString());
                        done = _cwdata.linesectionSetRecordCrossReference(cwid, fid);
                    }
                }
                catch (Exception err)
                {
                    string M = err.Message;
                }
            return false;
        }


        //-----------------------------------------------------------------------------------------------
        // This function creates the cross reference between the ArcMAP feature ID (ObjectID) and
        // our CompWorks ID;
        // Dbf Boring Locations table
        //-----------------------------------------------------------------------------------------------
        public bool boringlocationsUpdateCrossReference()
        {
            string layername = _layers.LayerName(LayerSystemUse.suBPBoringLocations);
            if (layername.Length == 0) return false;
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.SubFields = _layers.AttributeIndexName(layername) + ", " + pTableOwner.OIDFieldName;

            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            int CWFieldIndex = pTableOwner.FindField(_layers.AttributeIndexName(layername));
            int FeatureFieldIndex = pTableOwner.FindField(pTableOwner.OIDFieldName);
            bool done = false;
            using (ComReleaser comReleaser = new ComReleaser())
                try
                {
                    ICursor cursor = pTableOwner.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IRow row = null;
                    while ((row = cursor.NextRow()) != null)
                    {
                        string cwid = row.get_Value(CWFieldIndex).ToString();
                        int fid = Convert.ToInt32(row.get_Value(FeatureFieldIndex).ToString());
                        done = _cwdata.blSetRecordCrossReference(cwid, fid);
                    }
                }
                catch (Exception err)
                {
                    string M = err.Message;
                }
            return false;
        }


        //-----------------------------------------------------------------------------------------------
        // This function creates the cross reference between the ArcMAP feature ID (ObjectID) and
        // our CompWorks ID;
        // Dbf Corrosion Probe table
        //-----------------------------------------------------------------------------------------------
        public bool corrosionprobesUpdateCrossReference()
        {
            string layername = _layers.LayerName(LayerSystemUse.suBPCorrosionProbes);
            if (layername.Length == 0) return false;
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.SubFields = _layers.AttributeIndexName(layername) + ", " + pTableOwner.OIDFieldName;

            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            int CWFieldIndex = pTableOwner.FindField(_layers.AttributeIndexName(layername));
            int FeatureFieldIndex = pTableOwner.FindField(pTableOwner.OIDFieldName);
            bool done = false;
            using (ComReleaser comReleaser = new ComReleaser())
                try
                {
                    ICursor cursor = pTableOwner.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IRow row = null;
                    while ((row = cursor.NextRow()) != null)
                    {
                        string cwid = row.get_Value(CWFieldIndex).ToString();
                        int fid = Convert.ToInt32(row.get_Value(FeatureFieldIndex).ToString());
                        done = _cwdata.cpSetRecordCrossReference(cwid, fid);
                    }
                }
                catch (Exception err)
                {
                    string M = err.Message;
                }
            return false;
        }

        //-----------------------------------------------------------------------------------------------
        // This function creates the cross reference between the ArcMAP feature ID (ObjectID) and
        // our CompWorks ID;
        // Dbf Rectifiers table
        //-----------------------------------------------------------------------------------------------
        public bool rectifierUpdateCrossReference()
        {
            string layername = _layers.LayerName(LayerSystemUse.suBPRectifiers);
            if (layername.Length == 0) return false;
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.SubFields = _layers.AttributeIndexName(layername) + ", " + pTableOwner.OIDFieldName;

            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            int CWFieldIndex = pTableOwner.FindField(_layers.AttributeIndexName(layername));
            int FeatureFieldIndex = pTableOwner.FindField(pTableOwner.OIDFieldName);
            bool done = false;
            using (ComReleaser comReleaser = new ComReleaser())
                try
                {
                    ICursor cursor = pTableOwner.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IRow row = null;
                    while ((row = cursor.NextRow()) != null)
                    {
                        string cwid = row.get_Value(CWFieldIndex).ToString();
                        int fid = Convert.ToInt32(row.get_Value(FeatureFieldIndex).ToString());
                        done = _cwdata.rectSetRecordCrossReference(cwid, fid);
                    }
                }
                catch (Exception err)
                {
                    string M = err.Message;
                }
            return false;
        }

        //-----------------------------------------------------------------------------------------------
        // This function creates the cross reference between the ArcMAP feature ID (ObjectID) and
        // our CompWorks ID;
        // Dbf Stations table
        //-----------------------------------------------------------------------------------------------
        public bool teststationUpdateCrossReference()
        {
            string layername = _layers.LayerName(LayerSystemUse.suBPTestStations);
            if (layername.Length == 0) return false;
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.SubFields = _layers.AttributeIndexName(layername) + ", " + pTableOwner.OIDFieldName;

            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            int CWFieldIndex = pTableOwner.FindField(_layers.AttributeIndexName(layername));
            int FeatureFieldIndex = pTableOwner.FindField(pTableOwner.OIDFieldName);
            bool done = false;
            using (ComReleaser comReleaser = new ComReleaser())
                try
                {
                    ICursor cursor = pTableOwner.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IRow row = null;
                    while ((row = cursor.NextRow()) != null)
                    {
                        string cwid = row.get_Value(CWFieldIndex).ToString();
                        int fid = Convert.ToInt32(row.get_Value(FeatureFieldIndex).ToString());
                        done = _cwdata.rectSetRecordCrossReference(cwid, fid);
                    }
                }
                catch (Exception err)
                {
                    string M = err.Message;
                }
            return false;
        }

        //-----------------------------------------------------------------------------------------------
        // This function creates the cross reference between the ArcMAP feature ID (ObjectID) and
        // our CompWorks ID;
        //-----------------------------------------------------------------------------------------------
        public bool layerUpdateCrossReference(LayerSystemUse sysUse)
        {
            string layername = _layers.LayerName(sysUse);
            if (layername.Length == 0) return false;
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.SubFields = _layers.AttributeIndexName(layername) + ", " + pTableOwner.OIDFieldName;

            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            int CWFieldIndex = pTableOwner.FindField(_layers.AttributeIndexName(layername));
            int FeatureFieldIndex = pTableOwner.FindField(pTableOwner.OIDFieldName);
            bool done = false;
            using (ComReleaser comReleaser = new ComReleaser())
                try
                {
                    ICursor cursor = pTableOwner.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IRow row = null;
                    while ((row = cursor.NextRow()) != null)
                    {
                        string cwid = row.get_Value(CWFieldIndex).ToString();
                        int fid = Convert.ToInt32(row.get_Value(FeatureFieldIndex).ToString());
                        done = _cwdata.rectSetRecordCrossReference(cwid, fid);
                    }
                }
                catch (Exception err)
                {
                    string M = err.Message;
                }
            return false;
        }



        //-----------------------------------------------------------------------------------------------
        // This function updates the Risk Ranking for the Line Sections
        //-----------------------------------------------------------------------------------------------
        public bool linesectionUpdateRiskRanking()
        {
            string layername = _layers.LayerName(LayerSystemUse.suBPLineSections);
            if (layername.Length == 0) return false;
            //Open table
            IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
            //IFeature pBLID;
            ITable pTableOwner = pFeatureWorkspace.OpenTable(layername);
            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();
            queryFilter.SubFields = _layers.AttributeIndexName(layername) + ", " + pTableOwner.OIDFieldName;

            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            int CWFieldIndex = pTableOwner.FindField(_layers.AttributeIndexName(layername));
            int FeatureFieldIndex = pTableOwner.FindField(pTableOwner.OIDFieldName);
            bool done = false;
            using (ComReleaser comReleaser = new ComReleaser())
                try
                {
                    ICursor cursor = pTableOwner.Search(queryFilter, true);
                    comReleaser.ManageLifetime(cursor);
                    IRow row = null;
                    while ((row = cursor.NextRow()) != null)
                    {
                        int cwid = Convert.ToInt32(row.get_Value(CWFieldIndex).ToString());
                        int fid = Convert.ToInt32(row.get_Value(FeatureFieldIndex).ToString());
                        done = _cwdata.linesectionSetRecordCrossReference(cwid, fid);
                    }
                }
                catch (Exception err)
                {
                    string M = err.Message;
                }
            return false;
        }



        private void linesectionGetRankingStatistics(IFeatureLayer m_featureLayer, LayerDef ldef, string riskFieldName, ref double minriskvalue, ref double maxriskvalue)
        {

            //Find the selected field in the feature layer
            IFeatureClass featureClass = m_featureLayer.FeatureClass;
            IField field = featureClass.Fields.get_Field(featureClass.FindField(riskFieldName));

            //Get a feature cursor
            ICursor cursor = (ICursor)m_featureLayer.Search(null, false);

            //Create a DataStatistics object and initialize properties
            IDataStatistics dataStatistics = new DataStatisticsClass();
            dataStatistics.Field = field.Name;
            dataStatistics.Cursor = cursor;

            //Get the result statistics
            IStatisticsResults statisticsResults = dataStatistics.Statistics;

            //Set the values min and max values
            minriskvalue = statisticsResults.Minimum;
            maxriskvalue = statisticsResults.Maximum;
        }


        //-----------------------------------------------------------------------------------------------
        // Sets the Line Section Layer Renderer to Color by Risk Ranking
        //-----------------------------------------------------------------------------------------------
        public bool linesectionRenderByRiskRanking(string fieldNameToUse)
        {
            double minriskvalue = 0;
            double maxriskvalue = 0;
            int i = 0;
            LayerDef ldef = new LayerDef();
            _layers.GetLayerDef(LayerSystemUse.suBPLineSections, ref ldef);
            IGeoFeatureLayer geofeaturelayer = null;

            // Check to see if the field to use exists
            if (!LayerAttributeFieldExists(ldef.layername, fieldNameToUse))
                return false;


            IMap map = _activeview.FocusMap;
            string ln = "";
            for (i = 0; i <= map.LayerCount - 1; i++)
            {
                ln = map.get_Layer(i).Name;
                if (ln == ldef.layername)
                {
                    geofeaturelayer = (IGeoFeatureLayer)map.get_Layer(i);
                    break;
                }
            }
            IFeatureLayer m_featureLayer;
            if (geofeaturelayer == null) return false;
            m_featureLayer = geofeaturelayer;

            // Get Statistics Values for the Selected Field
            linesectionGetRankingStatistics(m_featureLayer, ldef, fieldNameToUse, ref minriskvalue, ref maxriskvalue);
            if (minriskvalue >= maxriskvalue) return false;
            List<RiskRankingRange> pRiskRankingRanges = new List<RiskRankingRange>();
            if (fieldNameToUse == "ConRanking")
            {
                if (_cwdata.RiskRankingRanges != null)
                    pRiskRankingRanges = _cwdata.RiskRankingRangesCons;
            }
            else
                if (fieldNameToUse == "SusRanking")
                {
                    pRiskRankingRanges = _cwdata.RiskRankingRangesSusc;
                }
                else
                    pRiskRankingRanges = _cwdata.RiskRankingRanges;

            if (pRiskRankingRanges.Count == 0) return false;
            if (m_featureLayer == null) return false;

            ClassBreaksRenderer m_classBreaksRenderer = new ClassBreaksRendererClass();
            m_classBreaksRenderer.Field = fieldNameToUse;
            m_classBreaksRenderer.BreakCount = pRiskRankingRanges.Count;
            m_classBreaksRenderer.MinimumBreak = 0;
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
            ldef.m_classBreaksRenderer = m_classBreaksRenderer;
            //Set the new renderer 
            if (ldef.m_classBreaksRenderer != null)
            {
                geofeaturelayer.Renderer = (IFeatureRenderer)ldef.m_classBreaksRenderer;
                //Trigger contents changed event for TOCControl
                _activeview.ContentsChanged();
                //Refresh the display 
                _activeview.PartialRefresh(esriViewDrawPhase.esriViewGeography, geofeaturelayer, null);
                return SetLayerRenderer(ldef);
            }
            else
                return false;
        }


        //-----------------------------------------------------------------------------------------------
        // Sets the Layer's Renderer to Basic Line
        //-----------------------------------------------------------------------------------------------
        public void linesection_AssignDefaultRenderer()
        {
            LayerDef ldef = new LayerDef();
            _layers.GetLayerDef(LayerSystemUse.suBPLineSections, ref ldef);
            SimpleRenderer m_SimpleRenderer = new SimpleRendererClass();
            IRgbColor rgbColor = new RgbColorClass();
            rgbColor.Blue = 143;
            rgbColor.Red = 0;
            rgbColor.Green = 57;
            SimpleLineSymbolClass simpleLineSymbol = new SimpleLineSymbolClass();
            simpleLineSymbol.Color = rgbColor;
            m_SimpleRenderer.Symbol = (ISymbol)simpleLineSymbol;
            ldef.m_SimpleRenderer = m_SimpleRenderer;
            SetLayerRendererBasicLine(ldef);
        }


        //-----------------------------------------------------------------------------------------------
        // Update Line Section Risk Ranking Values in Attributes table
        //-----------------------------------------------------------------------------------------------
        public void linesectionUpdateRR(int zone_ID, double overall_RR, double susc_RR, double cons_RR)
        {
            try
            {

                LayerDef ldef = new LayerDef();
                _layers.GetLayerDef(LayerSystemUse.suBPLineSections, ref ldef);
                IFeatureWorkspace pFeatureWorkspace = _workspace as IFeatureWorkspace;
                IFeatureClass featureClass = pFeatureWorkspace.OpenFeatureClass(ldef.layername);

                // Restrict the number of features to be updated.
                IQueryFilter queryFilter = new QueryFilterClass();
                queryFilter.WhereClause = string.Format("{0} = {1}", ldef.GIS_Attribute_Index_Field, zone_ID);
                queryFilter.SubFields = "Overall_Ranking, SusRanking, ConRanking";

                // Use IFeatureClass.Update to populate IFeatureCursor.
                IFeatureCursor updateCursor = featureClass.Update(queryFilter, false);

                int overall_idx = featureClass.FindField("Overall_Ranking");
                int susc_idx = featureClass.FindField("SusRanking");
                int cons_idx = featureClass.FindField("ConRanking");
                IFeature feature = null;

                while ((feature = updateCursor.NextFeature()) != null)
                {
                    feature.set_Value(overall_idx, overall_RR);
                    feature.set_Value(susc_idx, susc_RR);
                    feature.set_Value(cons_idx, cons_RR);
                    updateCursor.UpdateFeature(feature);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(updateCursor);
            }
            catch (Exception ex)
            {
            }
        }


        //--------------------------------------------------------------------------
        // Opens the Line Section Details form for the specified ID
        //--------------------------------------------------------------------------
        public void OpenForm_LineSectionDetails(int ID)
        {
            BPLineSectionDetails detailform = new BPLineSectionDetails(ID, _cwdata._connectionstring, true);
            detailform.ShowDialog();
            detailform.Dispose();
        }

        //--------------------------------------------------------------------------
        // Opens the Line Section Inspections Details form for the specified ID
        //--------------------------------------------------------------------------
        public void OpenForm_LineSectionInspections(int ID)
        {
            BPLineSectionInspections detailform = new BPLineSectionInspections(ID, _cwdata._connectionstring);
            detailform.ShowDialog();
            detailform.Dispose();
        }

        //--------------------------------------------------------------------------
        // Opens the Line Details form for the specified ID
        //--------------------------------------------------------------------------
        public void OpenForm_LineDetails(int ID)
        {
            BPLineDetails detailform = new BPLineDetails(ID, _cwdata._connectionstring, true);
            detailform.ShowDialog();
            detailform.Dispose();
        }

        //--------------------------------------------------------------------------
        // Opens the Boring Location Details form for the specified ID
        //--------------------------------------------------------------------------
        public void OpenForm_BoringLocationDetails(string ID)
        {
            BEDetails detailform = new BEDetails(ID, _cwdata._connectionstring, true);
            detailform.ShowDialog();
            detailform.Dispose();
        }

        //--------------------------------------------------------------------------
        // Opens the Corrosion Probe Details form for the specified ID
        //--------------------------------------------------------------------------
        public void OpenForm_CorrosionProbeDetails(string ID)
        {
            BPCP_Details detailform = new BPCP_Details(ID, _cwdata._connectionstring);
            detailform.ShowDialog();
            detailform.Dispose();
        }

        //--------------------------------------------------------------------------
        // Opens the Rectifier Details form for the specified ID
        //--------------------------------------------------------------------------
        public void OpenForm_RectifierDetails(int ID)
        {
            BPRectifierDetails detailform = new BPRectifierDetails(ID, _cwdata._connectionstring);
            detailform.ShowDialog();
            detailform.Dispose();
        }

        //--------------------------------------------------------------------------
        // Opens the Test Station Details form for the specified ID
        //--------------------------------------------------------------------------
        public void OpenForm_TestStationDetails(int ID)
        {
            BPTestStationDetails detailform = new BPTestStationDetails(ID, _cwdata._connectionstring);
            detailform.ShowDialog();
            detailform.Dispose();
        }


        //-----------------------------------------------------------------------------------------------
        // Clears the selection set for all layers
        //-----------------------------------------------------------------------------------------------
        public void select_ClearAllSelections()
        {
            LayerSystemUse[] layerstoclear = { LayerSystemUse.suBPLines, LayerSystemUse.suBPLineSections, LayerSystemUse.suBPBoringLocations, LayerSystemUse.suBPCorrosionProbes };
            select_ClearLayerSelections(layerstoclear);
        }


        //-----------------------------------------------------------------------------------------------
        // Check Database for any required Field Updates
        //-----------------------------------------------------------------------------------------------
        public bool CheckForFieldUpdates()
        {
            return (_cwdata.linesectionCheckForUpdates("R"));
        }

        //--------------------------------------------------------------------------
        // Process a detail form for a Layer
        //--------------------------------------------------------------------------
        public void ShowDetailForm(SelectedItemData selItem)
        {
            switch (_layers.SystemUse(selItem.layername))
            {
                case (LayerSystemUse.suBPLines):
                    {
                        OpenForm_LineDetails(Convert.ToInt32(selItem.CWKeyValue));
                        break;
                    }
                case (LayerSystemUse.suBPLineSections):
                    {
                        OpenForm_LineSectionDetails(Convert.ToInt32(selItem.CWKeyValue));
                        break;
                    }
                case (LayerSystemUse.suBPBoringLocations):
                    {
                        OpenForm_BoringLocationDetails(selItem.CWKeyValue);
                        break;
                    }
                case (LayerSystemUse.suBPCorrosionProbes):
                    {
                        OpenForm_CorrosionProbeDetails(selItem.CWKeyValue);
                        break;
                    }
                case (LayerSystemUse.suBPRectifiers):
                    {
                        OpenForm_RectifierDetails(Convert.ToInt32(selItem.CWKeyValue));
                        break;
                    }
                case (LayerSystemUse.suBPTestStations):
                    {
                        OpenForm_TestStationDetails(Convert.ToInt32(selItem.CWKeyValue));
                        break;
                    }
                   
            } // switch
        }


        public void CreateFeature(IFeatureClass featureClass, IPoint point)
        {
            // Ensure the feature class contains points.
            if (featureClass.ShapeType != esriGeometryType.esriGeometryPoint)
            {
                return;
            }

            // Build the feature.
            IFeature feature = featureClass.CreateFeature();
            feature.Shape = point;

            // Apply the appropriate subtype to the feature.
            ISubtypes subtypes = (ISubtypes)featureClass;
            IRowSubtypes rowSubtypes = (IRowSubtypes)feature;
            if (subtypes.HasSubtype)
            {
                // In this example, the value of 3 represents the Cross subtype.
                rowSubtypes.SubtypeCode = 3;
            }

            // Initialize any default values the feature has.
            rowSubtypes.InitDefaultValues();

            // Update the value on a string field that indicates who installed the feature.
            int contractorFieldIndex = featureClass.FindField("CONTRACTOR");
            feature.set_Value(contractorFieldIndex, "Test111");

            // Commit the new feature to the geodatabase.
            feature.Store();
        }        


    }
}
