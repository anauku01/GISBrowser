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



    public class ArcEngineSW : ArcEngineClass
    {

        //-----------------------------------------------------------------------------------------------
        // class constructor
        // pass the mx datafile to open
        //-----------------------------------------------------------------------------------------------
        public ArcEngineSW(string cwdatafile, string cwsystemfile)
            : base(cwdatafile, cwsystemfile)
        {
            //
        }


        //--------------------------------------------------------------------------
        // Opens the Line Details form for the specified ID
        //--------------------------------------------------------------------------
        public void OpenForm_LineDetails(int ID)
        {
            SWLineDetails detailform = new SWLineDetails(ID, _cwdata._connectionstring, true);
            detailform.ShowDialog();
            detailform.Dispose();
        }


        //--------------------------------------------------------------------------
        // Process a detail form for a Layer
        //--------------------------------------------------------------------------
        public void ShowDetailForm(SelectedItemData selItem)
        {
            switch (_layers.SystemUse(selItem.layername))
            {
                case (LayerSystemUse.suSWLines):
                    {
                        OpenForm_LineDetails(Convert.ToInt32(selItem.CWKeyValue));
                        break;
                    }

            } // switch
        }



        //-----------------------------------------------------------------------------------------------
        // This function creates the cross reference between the ArcMAP feature ID (ObjectID) and
        // our CompWorks ID;
        //-----------------------------------------------------------------------------------------------
        public bool UpdateCrossReference(string layerName)
        {
            string layername = layerName;
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



    }

}