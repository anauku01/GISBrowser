using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.ADF;
using ESRI.ArcGIS.ADF.Connection.Local;
using ESRI.ArcGIS;
using ESRI.ArcGIS.Geometry;


namespace COMPGIS1
{
    public class QueryMap
    {
        public QueryMap()
        {
            //Open workspace for the File GDB
            IWorkspace pExampleWorkspace = null;
            string path = @"C:\GISData\LaSalleDataXY.gdb";
            pExampleWorkspace = FileGDBWorkspaceFromPath(path);

            if (pExampleWorkspace != null)
                Console.WriteLine("Successfully open FileGDB File!\n");

            try
            {
                //Perform Query on a Sample Table
                //PerformTableQuery(pExampleWorkspace);

                //Perform Spatial Query
                PerformSpatialQuery(pExampleWorkspace);
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }

        }

    //    //public void QM();
   
    //    private static LicenseInitializer m_AOLicenseInitializer = new LicenseInitializer();

        private static IWorkspace FileGDBWorkspaceFromPath(String path)
        {
            IWorkspaceFactory workspaceFactory = new FileGDBWorkspaceFactory();
            return workspaceFactory.OpenFromFile(path, 0);
        }

        private static void PerformTableQuery(IWorkspace pWorkspace)
        {
            //Open table
            IFeatureWorkspace pFeatureWorkspace = pWorkspace as IFeatureWorkspace;
            ITable pTableOwner = pFeatureWorkspace.OpenTable("BoringLocations");


            // Create the query filter.
            IQueryFilter queryFilter = new QueryFilterClass();

            // Select the fields to be returned—the name and address of the businesses.
            queryFilter.SubFields = "OBJECTID, LOCATION, POINT_X, POINT_Y";

            // Set the filter to return only restaurants.
            queryFilter.WhereClause = "OBJECTID>5 ";

            // Use the PostfixClause to alphabetically order the set by name.
            IQueryFilterDefinition queryFilterDef = (IQueryFilterDefinition)queryFilter;
            queryFilterDef.PostfixClause = "ORDER BY OBJECTID";

            // Output the returned names and addresses.
            int ObjectIndex = pTableOwner.FindField("OBJECTID");
            int LocationIndex = pTableOwner.FindField("LOCATION");
            using (ComReleaser comReleaser = new ComReleaser())
            {
                ICursor cursor = pTableOwner.Search(queryFilter, true);

                comReleaser.ManageLifetime(cursor);
                IRow row = null;
                while ((row = cursor.NextRow()) != null)
                {
                    String Object = Convert.ToString(row.get_Value(ObjectIndex));
                    String Location = Convert.ToString(row.get_Value(LocationIndex));
                    Console.WriteLine("{0} - {1}", Object, Location);
                    System.Diagnostics.Debug.WriteLine("{0} - {1}", Object, Location);
                    Console.Read();
                }
            }
        }


        private static void PerformSpatialQuery(IWorkspace pWorkspace)
        {
            //Open FeatureClass
            IFeatureWorkspace pFeatureWorkspace = pWorkspace as IFeatureWorkspace;
            IFeatureClass pBoringLocationsFeatureClass = pFeatureWorkspace.OpenFeatureClass("BoringLocations");
            IFeatureClass pLineSectionFeatureClass = pFeatureWorkspace.OpenFeatureClass("Line_Sections");
            IFeature blockFeature = pBoringLocationsFeatureClass.GetFeature(5);
            IGeometry queryGeometry = blockFeature.ShapeCopy;
            // Buffer all the selected features by the buffer distance and create a new polygon element from each result
            ESRI.ArcGIS.Geometry.ITopologicalOperator topologicalOperator;
            ESRI.ArcGIS.Carto.IElement element=null;
            //while (!(blockFeature == null))
            //{
                topologicalOperator = (ESRI.ArcGIS.Geometry.ITopologicalOperator)blockFeature.Shape; // Explicit Cast
                element = new ESRI.ArcGIS.Carto.PolygonElementClass();
                element.Geometry = topologicalOperator.Buffer(20);
                //graphicsContainer.AddElement(element, 0);
                //blockFeature = enumFeature.Next();
            //}
            // Create the query filter.
            ISpatialFilter queryFilter = new SpatialFilterClass()
            {

                Geometry = element.Geometry,
                GeometryField = pBoringLocationsFeatureClass.ShapeFieldName,
                SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            };

            // Output the returned names and addresses.
            int nameRes = pLineSectionFeatureClass.FindField("Line_Secti");
            using (ComReleaser comReleaser = new ComReleaser())
            {
                IFeatureCursor cursor = pLineSectionFeatureClass.Search(queryFilter, true);
                comReleaser.ManageLifetime(cursor);
                IFeature feature = null;
                while ((feature = cursor.NextFeature()) != null)
                {
                    String name = Convert.ToString(feature.get_Value(nameRes));
                    System.Diagnostics.Debug.WriteLine("{0} - {1:f}:{2:f}",
                        name, feature.Shape.Envelope.XMax, feature.Shape.Envelope.YMax);
                }
            }
        }


    //    [STAThread()]
    //    static void Main(string[] args)
    //    {
    //        //ESRI License Initializer generated code.
    //        m_AOLicenseInitializer.InitializeApplication(new esriLicenseProductCode[] { esriLicenseProductCode.esriLicenseProductCodeBasic, esriLicenseProductCode.esriLicenseProductCodeStandard, esriLicenseProductCode.esriLicenseProductCodeAdvanced
    //        },
    //        new esriLicenseExtensionCode[] { });

    //        //Open workspace for the File GDB
    //        IWorkspace pExampleWorkspace = null;
    //        string path = @"C:\GISData\LaSalleData.gdb";
    //        pExampleWorkspace = FileGDBWorkspaceFromPath(path);

    //        if (pExampleWorkspace != null)
    //            Console.WriteLine("Successfully open FileGDB File!\n");

    //        try
    //        {
    //            //Perform Query on a Sample Table
    //            //PerformTableQuery(pExampleWorkspace);

    //            //Perform Spatial Query
    //            PerformSpatialQuery(pExampleWorkspace);
    //        }
    //        catch (Exception err)
    //        {
    //            Console.WriteLine(err.Message);
    //        }

    //        //ESRI License Initializer generated code.
    //        //Do not make any call to ArcObjects after ShutDownApplication()
    //        m_AOLicenseInitializer.ShutdownApplication();
    //    }
    }
}

