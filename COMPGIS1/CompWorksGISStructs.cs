using System.Collections.Generic;
using System;
using System.Drawing;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Display;

namespace COMPGIS1
{

    public enum CompWorksApps {cwNone,cwBuriedPiping,cwServiceWater,cwFAC};
    public enum LayerSystemUse {suNone,suBPLines,suBPLineSections,suBPBoringLocations,suBPCorrosionProbes,suBPRectifiers,suBPTestStations,suBPMonitoringWell, // 0 - 7

                                suSWLines,suSWFlowSegments,suSWComponents}; // 8 - 10

    public class LayerDef
    {        
        public string layername;
        public string Description;
        public string BaseFilterQuery;
        public bool LayerVisibleAtStartup;
        public LayerSystemUse SystemUse;
        public string GIS_CW_Index_Field;  // The layer's CompWorks Index field name
        public string GIS_CW_Table_Name; // The layer's CompWorks table name
        public string GIS_CW_Feature_Display_Field; // The CompWorks identifier field name to show in tree/list views etc...
        public bool GIS_Key_Field_Is_String; // True if the CW Key field is a string (i.e. Boring Locations, Corrosion Probes)
        public string GIS_Attribute_Index_Field; // The CompWorks Keyfield name in the layer attribute table; Loaded at startup
        public string GIS_Attribute_Table_Object_Index_Field; // The primary key field name for the feature's attribute table
        public int GISLayerID;

        public string LayerTableName = "";                                       

        public List<SQLFilterClass> SavedFilters;
        public int FilterCount
        { get { return SavedFilters.Count; }  }

        private SQLFilterClass _activelayer = null;
        public SQLFilterClass ActiveFilter { get { return _activelayer; } set { _activelayer = value; } }

        public ClassBreaksRenderer m_classBreaksRenderer = null;
        public SimpleRenderer m_SimpleRenderer = null;

        public int DisplayMode = 0;
        public int DisplayByField = 0;



        //----------------------------------
        // Constructor
        //----------------------------------
        public LayerDef()
        {
            SavedFilters = new List<SQLFilterClass>();
        }

        //----------------------------------
        // Add a Saved Filter
        //----------------------------------
        public bool AddSavedFilter(SQLFilterClass newfilter)
        {
            if (newfilter != null)
            {
                SavedFilters.Add(newfilter);
                return true;
            }
            return false;
        }

        //----------------------------------
        //  Returns the saved filter by FilterName
        //----------------------------------
        public SQLFilterClass GetSavedFilter(string filtername)
        {
            int i;
            for (i = 0; i <= (SavedFilters.Count - 1); i++)
            {
                if (SavedFilters[i].FilterName == filtername)
                {
                    return SavedFilters[i];
                }
            }
            return null;
        }

        //----------------------------------
        //  Removes a saved filter by ID
        //----------------------------------
        public bool RemoveSavedFilter(int ID)
        {
            int i;
            for (i = 0; i <= (SavedFilters.Count - 1); i++)
            {
                if (SavedFilters[i].ID == ID)
                {
                    SavedFilters.RemoveAt(i);
                    return true;
                }
            }
            return false;
        }


        //----------------------------------
        //  Returns the saved filter by ordinal index
        //----------------------------------
        public SQLFilterClass GetSavedFilter(int index)
        {
            if ((index < 0) || (index > (SavedFilters.Count - 1))) return null;
            return SavedFilters[index];
        }


    }


    //..........................................................................................................................................
    //------------------------------------------------
    // LayerDefs class    
    //------------------------------------------------
    public class LayerDefs
    {
        //................
        // Properties
        //................
        private List<LayerDef> _layers;

        private const int su_layerLines = 1;
        private const int su_layerLineSections = 2;
        private const int su_layerBoringLocations = 3;
        private const int su_layerCorrosionProbes = 4;


        private const int _layernametypescount = 11;
        private string[] _layernametypes = new string[_layernametypescount] { "None", "Lines", "Line Sections", "Boring Locations", "Corrosion Probes", "Rectifiers", "Test Stations", "Monitoring Well", "Lines", "Flow Segments", "Components" };

        public int Count
        {
            get { return _layers.Count; }
        }

        //................
        // Methods
        //................

        //----------------------------------
        // Constructor
        //----------------------------------
        public LayerDefs()
        {
            _layers = new List<LayerDef>();
        }

        //----------------------------------
        // Return the Layer name from the
        // system_use constant value
        //----------------------------------
        public string LayerName(LayerSystemUse system_use)
        {
            int i = 0;
            for (i=0;i<=(_layers.Count-1);i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    return _layers[i].layername;
                }
            }
            return "";
        }

        //----------------------------------
        // Return the CW Index Field name from the
        // system_use constant value
        //----------------------------------
        public string CWIndexName(LayerSystemUse system_use)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    return _layers[i].GIS_CW_Index_Field;
                }
            }
            return "";
        }

        //----------------------------------
        // Return the CW Index Field name from the
        // system_use constant value
        //----------------------------------
        public string CWIndexName(string layername)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (string.Compare(_layers[i].layername,layername)==0)
                {
                    return _layers[i].GIS_CW_Index_Field;
                }
            }
            return "";
        }

        //----------------------------------
        // Return boolean indicating if the key field for the layer is a string
        // using the layer name
        // Overload 1
        //----------------------------------
        public bool KeyFieldIsString(string layername)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (string.Compare(_layers[i].layername, layername) == 0)
                {
                    return _layers[i].GIS_Key_Field_Is_String;
                }
            }
            return false;
        }

        //----------------------------------
        // Return boolean indicating if the key field for the layer is a string
        // using the System_Use 
        // Overload 2
        //----------------------------------
        public bool KeyFieldIsString(LayerSystemUse system_use)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    return _layers[i].GIS_Key_Field_Is_String;
                }
            }
            return false;
        }

        //----------------------------------
        // Return string with the Table Name For the Layer
        // using the System_Use 
        // Overload 1
        //----------------------------------
        public string CWTableName(LayerSystemUse system_use)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    return _layers[i].GIS_CW_Table_Name;
                }
            }
            return "";
        }

        //----------------------------------
        // Return string with the Table Name For the Layer
        // using the LayerName 
        // Overload 2
        //----------------------------------
        public string CWTableName(string layername)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (string.Compare(_layers[i].layername, layername) == 0)
                {
                    return _layers[i].GIS_CW_Table_Name;
                }
            }
            return "";
        }


        //----------------------------------
        // Return the Name of the Feature's Index Field (ObjectIDxxx)
        // Overload 1
        //----------------------------------
        public string AttributeObjectIndexName(LayerSystemUse system_use)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    return _layers[i].GIS_Attribute_Table_Object_Index_Field;
                }
            }
            return "";
        }


        //----------------------------------
        // Return the N of the Feature's Index Field (ObjectIDxxx)
        // Overload 2
        //----------------------------------
        public string AttributeObjectIndexName(string layername)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (string.Compare(_layers[i].layername, layername) == 0)
                {
                    return _layers[i].GIS_Attribute_Table_Object_Index_Field;
                }
            }
            return "";
        }


        //----------------------------------
        // Return the Attribute (ArcGIS Feature) CW Index Field name from the
        // system_use constant value
        // Overload 1
        //----------------------------------
        public string AttributeIndexName(LayerSystemUse system_use)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    return _layers[i].GIS_Attribute_Index_Field;
                }
            }
            return "";
        }


        //----------------------------------
        // Return the Attribute (ArcGIS Feature) CW Index Field name from the
        // layer name
        // Overload 2
        //----------------------------------
        public string AttributeIndexName(string layername)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (string.Compare(_layers[i].layername, layername) == 0)
                {
                    return _layers[i].GIS_Attribute_Index_Field;
                }
            }
            return "";
        }


        //----------------------------------
        // Return the Base Filter Query for the layer based upon
        // system_use constant value
        // Overload 1
        //----------------------------------
        public string BaseFilterQuery(LayerSystemUse system_use)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    return _layers[i].BaseFilterQuery;
                }
            }
            return "";
        }


        //----------------------------------
        // Return the Base Filter Query for the layer based upon
        // layer name
        // Overload 2
        //----------------------------------
        public string BaseFilterQuery(string layername)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (string.Compare(_layers[i].layername, layername) == 0)
                {
                    return _layers[i].BaseFilterQuery;
                }
            }
            return "";
        }


        //----------------------------------
        // Return the SystemUse value from
        // list ordinal index value
        //----------------------------------
        public LayerSystemUse SystemUseOrdinal(int index)
        {
            if ((index < 0) || (index > (_layers.Count - 1))) return LayerSystemUse.suNone;
            return _layers[index].SystemUse;
        }


        //----------------------------------
        // Return the SystemUse value from
        // layer name
        //----------------------------------
        public LayerSystemUse SystemUse(string layername)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (string.Compare(_layers[i].layername, layername) == 0)
                {
                    return _layers[i].SystemUse;
                }
            }
            return 0;
        }


        //----------------------------------
        // Return string with the Feature Display Field Name For the Layer
        // using the System_Use 
        // Overload 1
        //----------------------------------
        public string CWFeatureDisplayField(LayerSystemUse system_use)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    return _layers[i].GIS_CW_Feature_Display_Field;
                }
            }
            return "";
        }


        //----------------------------------
        // Return string with the Feature Display Field Name For the Layer
        // using the LayerName 
        // Overload 2
        //----------------------------------
        public string CWFeatureDisplayField(string layername)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (string.Compare(_layers[i].layername, layername) == 0)
                {
                    return _layers[i].GIS_CW_Feature_Display_Field;
                }
            }
            return "";
        }

        //----------------------------------
        // Return the Layer class from the
        // ordinal/index value
        //----------------------------------
        public bool GetLayerDefOrdinal(int index, ref LayerDef LyrDef)
        {
            int i = 0;
            if ((i>=0) && (index <= (_layers.Count-1))) 
            {
                LyrDef = _layers[index];
                return true;
            }
            else
                return false;
        }

        //----------------------------------
        // Return the Layer class from the
        // system_use constant value
        //----------------------------------
        public bool GetLayerDef(LayerSystemUse system_use, ref LayerDef LyrDef)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    LyrDef = _layers[i];
                    return true;
                }
            }
            return false;
        }

        //----------------------------------
        // Return the Layer class from the
        // layername
        //----------------------------------
        public bool GetLayerDef(string layername, ref LayerDef LyrDef)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].layername == layername)
                {
                    LyrDef = _layers[i];
                    return true;
                }
            }
            return false;
        }


        //----------------------------------
        // Returns true if the system is a system layer
        //----------------------------------
        public bool IsSystemLayer(string layername)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (string.Compare(_layers[i].layername, layername) == 0)
                {
                    return true;
                }
            }
            return false;
        }


        //----------------------------------
        // Clear the ActiveFilter for all layers
        //----------------------------------
        public void ClearAllActiveFilters()
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
                _layers[i].ActiveFilter = null;
        }


        //----------------------------------
        // Adds a filter to the specified Layer's list 
        // system_use constant value
        // NOTE: Does not save it to database
        //----------------------------------
        public bool AddFiltertoList(LayerSystemUse system_use, SQLFilterClass filter)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    return _layers[i].AddSavedFilter(filter);
                }
            }
            return false;
        }


        //----------------------------------
        // Return the ActiveFilter for a layer from
        // system_use constant value
        //----------------------------------
        public SQLFilterClass ActiveFilter(LayerSystemUse system_use)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    return _layers[i].ActiveFilter;
                }
            }
            return null;
        }

        //----------------------------------
        // Return the ActiveFilter for a layer from
        // system_use constant value
        //----------------------------------
        public SQLFilterClass ActiveFilter(string layername)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (string.Compare(_layers[i].layername, layername) == 0)
                {
                    return _layers[i].ActiveFilter;
                }
            }
            return null;
        }

        //----------------------------------
        // Clear the ActiveFilter for a layer from
        // system_use constant value
        //----------------------------------
        public bool ClearActiveFilter(LayerSystemUse system_use)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    _layers[i].ActiveFilter = null;
                    return true;
                }
            }
            return false;
        }


        //----------------------------------
        // Remove the ActiveFilter for a layer from
        // system_use constant value
        //----------------------------------
        public bool RemoveActiveFilter(LayerSystemUse system_use)
        {
            int i = 0;
            int id = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    if (_layers[i].ActiveFilter != null)
                    {
                        id = _layers[i].ActiveFilter.ID;                                                
                        _layers[i].ActiveFilter = null;
                        _layers[i].RemoveSavedFilter(id);
                        return true;                       
                    }
                }
            }
            return false;
        }


        //----------------------------------
        // Set the ActiveFilter for a layer from
        // system_use constant value
        //----------------------------------
        public bool SetActiveFilter(LayerSystemUse system_use, SQLFilterClass filter)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (_layers[i].SystemUse == system_use)
                {
                    _layers[i].ActiveFilter = filter;
                    return true;
                }
            }
            return false;
        }

        //----------------------------------
        // Set the ActiveFilter for a layer from
        // system_use constant value
        //----------------------------------
        public bool SetActiveFilter(string layername,SQLFilterClass filter)
        {
            int i = 0;
            for (i = 0; i <= (_layers.Count - 1); i++)
            {
                if (string.Compare(_layers[i].layername, layername) == 0)
                {
                    _layers[i].ActiveFilter = filter;
                    return true;
                }
            }
            return false;
        }



        //----------------------------------
        // Add a layer struct to the array
        //----------------------------------
        public int AddLayer(LayerDef layerdef)
        {
            
            // Add the base query if it's a system layer 
            switch (layerdef.SystemUse)
            {
                case (LayerSystemUse.suBPLines):
                    {
                        layerdef.BaseFilterQuery = "SELECT DISTINCT [Dbf Lines].LineID, [Dbf Lines].MapFeatureID, [Dbf Lines].Line, [Dbf Lines].Unit, [Dbf Lines].LineNumber, [Dbf Lines].LineDescription, [Dbf Lines].LineClass, [Dbf Lines].LineGroup, [Dbf Lines].LinePhase, [Dbf Lines].LineSafetyGrade, [Dbf Lines].LineSize, [Dbf Lines].BoreClass, [Dbf Lines].Category, [Dbf Lines].Criteria, [Dbf Lines].BasisComments, [Dbf Lines].FailureConsequence, [Dbf Lines].OtherSpec, [Dbf Lines].Operation, [Dbf Lines].Velocity, [Dbf Lines].NSIAC, [Dbf Lines].LicenseRenewalCommitment, [Dbf Lines].CommitmentNumber, [Dbf Lines].LRCRemarks, [Dbf Lines].IsGWPILine, [Dbf Lines].[Level of Susceptability], [Dbf Lines].Temperature, [Dbf Setup -  Data].ValueData AS ['Line Content'], [Dbf Lines].[Operating Time], [Dbf Lines].EvalComments, [Dbf Lines].SNMComments, [Dbf Lines].References, [Dbf Lines].PlantExperience, [Dbf Lines].PlantExperienceDescription, [Dbf Lines].IndustryExperience, [Dbf Lines].IndustryExperienceDescription, [Dbf Lines].CheckWorksLineName, [Dbf Lines].CheckWorksLineNumber, [Dbf Lines].[Failure Affects Safe Shutdown or CDF], [Dbf Lines].[Radiological Content], [Dbf Lines].Inservice, [Dbf Lines].[Tank Type], [Dbf Lines].[Begin Point], [Dbf Lines].[End Point] FROM [Dbf Lines] LEFT JOIN [Dbf Setup -  Data] ON [Dbf Lines].[Line Content] = [Dbf Setup -  Data].SetupID";
                        break;
                    }
                case (LayerSystemUse.suBPLineSections):
                    {
                        layerdef.BaseFilterQuery = "SELECT [Dbf Line Section].[Zone ID], [Dbf Line Section].LineID, [Dbf Line Section].MapFeatureID, [Dbf Line Section].[Line Section], [Dbf Line Section].Unit, [Dbf Line Section].[Line Section Ref #], [Dbf Line Section].[Line Section Type], [Dbf Line Section].[Total Length], [Dbf Line Section].[Line Section Group], [Dbf Line Section].[Line Section Description], [Dbf Line Section].Node, [Dbf Line Section].Comments, [Dbf Line Section].[Date Install], [Dbf Line Section].Line_Content, [Dbf Line Section].ID_Leak_Likelihood, [Dbf Line Section].ID_Break_Likelihood, [Dbf Line Section].ID_Occlusion_Likelihood, [Dbf Line Section].OD_Leak_Likelihood, [Dbf Line Section].OD_Break_Likelihood, [Dbf Line Section].[ID Leak Cell], [Dbf Line Section].[ID Break Cell], [Dbf Line Section].[ID Occlusion Cell], [Dbf Line Section].[OD Leak Cell], [Dbf Line Section].[OD Break Cell], [Dbf Line Section].[ID All Cell], [Dbf Line Section].[OD All Cell], [Dbf Line Section].[ID and OD Cell], [Dbf Line Section].[ID Leak Cell Number], [Dbf Line Section].[ID Break Cell Number], [Dbf Line Section].[ID Occlusion Cell Number], [Dbf Line Section].[OD Leak Cell Number], [Dbf Line Section].[OD Break Cell Number], [Dbf Line Section].[ID All Cell Number], [Dbf Line Section].[OD All Cell Number], [Dbf Line Section].[ID OD Cell Number], [Dbf Line Section].LastModified, " +
                        "[Dbf Line Section].GUID, [Dbf Line Section].Notes1, [Dbf Line Section].Notes2, [Dbf Line Section].Notes3, [Dbf Line Section].Notes4, [Dbf Line Section].Notes5, [Dbf Line Section].Notes6, [Dbf Line Section].Notes7, [Dbf Line Section].isXML, [Dbf Line Section].Scheduled, [Dbf Line Section].ID, [Dbf Line Section].Custom1, [Dbf Line Section].Custom2, [Dbf Line Section].Custom3, [Dbf Line Section].Custom4, [Dbf Line Section].Custom5, [Dbf Line Section].Custom6, [Dbf Line Section].Custom7, [Dbf Line Section].Custom8, [Dbf Line Section].Custom9, [Dbf Line Section].[Contamination Level], [Dbf Line Section].Contaminated, [Dbf Line Section].[Contamination Type], [Dbf Line Section].[Dose Rate], [Dbf Line Section].Dose, [Dbf Line Section].[Zone Configuration], [Dbf Line Section].Coordinates, [Dbf Line Section].[Begin Point], [Dbf Line Section].[End Point], [Dbf Line Section].GUIDV2, [Dbf Line Section].Area, [Dbf Line Section].[Area Description], [Dbf Line Section].[ASME Class], [Dbf Line Section].Building, [Dbf Line Section].[Building Zone], [Dbf Line Section].Elevation, [Dbf Line Section].Location, [Dbf Line Section].[Location Comments], [Dbf Line Section].Plant, [Dbf Line Section].[Safety Related], [Dbf Line Section].Loop, [Dbf Line Section].Room, [Dbf Line Section].[Replacement Comments], [Dbf Line Section].[Scaffolding Req], "+
                        "[Dbf Line Section].[Scaffolding Type], [Dbf Line Section].[Scaffolding Date], [Dbf Line Section].[Insulation Removal], [Dbf Line Section].[Insulation Removal Type], [Dbf Line Section].[Insulation Date], [Dbf Line Section].Preparation, [Dbf Line Section].[Preparation Type], [Dbf Line Section].[Preparation Date], [Dbf Line Section].[Grid Required], [Dbf Line Section].[Grid Type], [Dbf Line Section].GridTemplateSN, [Dbf Line Section].AxialGridSize, [Dbf Line Section].RadialGridSize, [Dbf Line Section].[Exam Instructions], [Dbf Line Section].[Accessible Online], [Dbf Line Section].[Inspection Priority], [Dbf Line Section].[Inspection LocationID], [Dbf Line Section].[Initial Inspection Number], [Dbf Line Section].[Installation Spec], [Dbf Line Section].[Inspection Interval], [Dbf Line Section].[Replacement WorkOrder], [Dbf Line Section].PrevComponentID, [Dbf Line Section].[Reinspection Interval], [Dbf Line Section].ReinstallationOutage, [Dbf Line Section].NextScheduledInspection, [Dbf Line Section].DefaultInspectionMeth, [Dbf Line Section].[Repair-Replace], [Dbf Line Section].[Repair-Replace Date1], [Dbf Line Section].[Repair-Replace WO1], [Dbf Line Section].[Repair-Replace Date2], [Dbf Line Section].[Repair-Replace WO2], [Dbf Line Section].[Repair-Replace Date3], [Dbf Line Section].[Repair-Replace WO3], [Dbf Line Section].Outage1, "+
                        "[Dbf Line Section].Outage2, [Dbf Line Section].Outage3, [Dbf Line Section].[Date Install USExt], [Dbf Line Section].[ReinstallationOutage USExt], [Dbf Line Section].[Date Install MainDS], [Dbf Line Section].[ReinstallationOutage MainDS], [Dbf Line Section].[Date Install DSExt], [Dbf Line Section].[ReinstallationOutage DSExt], [Dbf Line Section].[Date Install Branch], [Dbf Line Section].[ReinstallationOutage Branch], [Dbf Line Section].[Date Install BranchExt], [Dbf Line Section].[ReinstallationOutage BranchExt], [Dbf Line Section].[Line Section Status], [Dbf Line Section].Insulated, [Dbf Line Section].[Corrective Actions], [Dbf Line Section].[Preventative Actions], [Dbf Line Section].[Priority Description], [Dbf Line Section].Review, [Dbf Line Section].[Reviewed By], [Dbf Line Section].[Reviewed Date], [Dbf Line Section].[Pressure Class], [Dbf Line Section].Material_Type, [Dbf Line Section].Schedule_Type, [Dbf Line Section].Potential, [Dbf Line Section].CompanyID, [Dbf Line Section].SubUnitID, [Dbf Line Section].SubSystemID, [Dbf Line Section].CompanyName, [Dbf Line Section].SubSystemName, [Dbf Line Section].LineName, [Dbf Line Section].SusRanking, [Dbf Line Section].ConRanking, [Dbf Line Section].Ranking, [Dbf Line Section].EngSusceptability, [Dbf Line Section].EngConsequence, [Dbf Line Section].EngJudgement, "+
                        "[Dbf Line Section].[Overall Ranking] FROM [Dbf Line Section]";
                        break;
                    }
                case (LayerSystemUse.suBPBoringLocations):
                    {
                        layerdef.BaseFilterQuery = "SELECT DISTINCT [Dbf Systems].System, [Dbf Lines].LineID, [Dbf Lines].Line, [Dbf Lines].Unit, [Dbf Lines].LineNumber, [Dbf Lines].LineDescription, [Dbf Lines].LineClass, [Dbf Lines].LineGroup, [Dbf Lines].LinePhase, [Dbf Lines].LineSafetyGrade, [Dbf Lines].LineSize, [Dbf Lines].BoreClass, [Dbf Lines].Category, [Dbf Lines].Criteria, [Dbf Lines].BasisComments, [Dbf Lines].FailureConsequence, [Dbf Lines].OtherSpec, [Dbf Lines].Operation, [Dbf Lines].Velocity, [Dbf Lines].NSIAC, [Dbf Lines].LicenseRenewalCommitment, [Dbf Lines].CommitmentNumber, [Dbf Lines].LRCRemarks, [Dbf Lines].IsGWPILine, [Dbf Lines].[Total Length], [Dbf Lines].[Level of Susceptability], [Dbf Lines].Temperature, [Dbf Lines].[Line Content], [Dbf Lines].[Operating Time], [Dbf Lines].EvalComments, [Dbf Lines].SNMComments, [Dbf Lines].References, [Dbf Lines].PlantExperience, [Dbf Lines].PlantExperienceDescription, [Dbf Lines].IndustryExperience, [Dbf Lines].IndustryExperienceDescription, [Dbf Lines].CheckWorksLineName, [Dbf Lines].CheckWorksLineNumber, [Dbf Lines].[Failure Affects Safe Shutdown or CDF], [Dbf Lines].[Radiological Content], [Dbf Lines].Inservice, [Dbf Lines].[Tank Type], [Dbf Lines].[Begin Point], [Dbf Lines].[End Point] FROM ([Dbf Lines] INNER JOIN [Dbf System_Line] ON [Dbf Lines].LineID = [Dbf System_Line].LineID) INNER JOIN [Dbf Systems] ON [Dbf System_Line].SystemID = [Dbf Systems].SystemID";
                        break;
                    }
                case (LayerSystemUse.suBPCorrosionProbes):
                    {
                        layerdef.BaseFilterQuery = "SELECT DISTINCT [Dbf Systems].System, [Dbf Lines].LineID, [Dbf Lines].Line, [Dbf Lines].Unit, [Dbf Lines].LineNumber, [Dbf Lines].LineDescription, [Dbf Lines].LineClass, [Dbf Lines].LineGroup, [Dbf Lines].LinePhase, [Dbf Lines].LineSafetyGrade, [Dbf Lines].LineSize, [Dbf Lines].BoreClass, [Dbf Lines].Category, [Dbf Lines].Criteria, [Dbf Lines].BasisComments, [Dbf Lines].FailureConsequence, [Dbf Lines].OtherSpec, [Dbf Lines].Operation, [Dbf Lines].Velocity, [Dbf Lines].NSIAC, [Dbf Lines].LicenseRenewalCommitment, [Dbf Lines].CommitmentNumber, [Dbf Lines].LRCRemarks, [Dbf Lines].IsGWPILine, [Dbf Lines].[Total Length], [Dbf Lines].[Level of Susceptability], [Dbf Lines].Temperature, [Dbf Lines].[Line Content], [Dbf Lines].[Operating Time], [Dbf Lines].EvalComments, [Dbf Lines].SNMComments, [Dbf Lines].References, [Dbf Lines].PlantExperience, [Dbf Lines].PlantExperienceDescription, [Dbf Lines].IndustryExperience, [Dbf Lines].IndustryExperienceDescription, [Dbf Lines].CheckWorksLineName, [Dbf Lines].CheckWorksLineNumber, [Dbf Lines].[Failure Affects Safe Shutdown or CDF], [Dbf Lines].[Radiological Content], [Dbf Lines].Inservice, [Dbf Lines].[Tank Type], [Dbf Lines].[Begin Point], [Dbf Lines].[End Point] FROM ([Dbf Lines] INNER JOIN [Dbf System_Line] ON [Dbf Lines].LineID = [Dbf System_Line].LineID) INNER JOIN [Dbf Systems] ON [Dbf System_Line].SystemID = [Dbf Systems].SystemID";
                        break;
                    }


                // -------------------------
                // Service Water 
                // -------------------------
                case (LayerSystemUse.suSWComponents):
                    {
                        layerdef.BaseFilterQuery = "SELECT [Dbf Components].[Item ID], [Dbf Components].[Summary No], [Dbf Components].[Component No], [Dbf Components].[Component No2], [Dbf Components].[Comp Description], [Dbf Components].[Component Type], [Dbf Systems].System, [Dbf Subsystems].SubSystem, [Dbf Lines].Line, [Dbf Zone].[Flow Segment], [Dbf Components].CompStatus, [Dbf Components].[Geometry Type], [Dbf Components].Configuration, [Dbf Components].[Comp Comments], [Dbf Components].[Comp Picture Path], [Dbf Components].[UPST Component], [Dbf Components].[Main Component], [Dbf Components].[DNST Component], [Dbf Components].Site, [Dbf Components].Unit, [Dbf Components].Building, [Dbf Components].Area, [Dbf Components].[Area Description], [Dbf Components].Location, [Dbf Components].[Location Comments], [Dbf Components].[Safety Related], [Dbf Components].Loop, [Dbf Components].[Building Zone], [Dbf Components].LineID, [Dbf Components].[ASME Class], [Dbf Components].[Zone ID], [Dbf Components].Room, [Dbf Components].Elevation, [Dbf Components].[Comp Elev], [Dbf Components].Coordinates, [Dbf Components].[Column No], [Dbf Components].[Row No], [Dbf Components].[Scaffolding Req], [Dbf Components].[Scaffolding Type], [Dbf Components].[Scaffolding Date], [Dbf Components].[Insulation Removal], [Dbf Components].[Insulation Removal Type], [Dbf Components].[Insulation Date], " +
                                                   "[Dbf Components].Preparation, [Dbf Components].[Preparation Type], [Dbf Components].[Preparation Date], [Dbf Components].[Grid Required], [Dbf Components].[Grid Type], [Dbf Components].GridTemplateSN, [Dbf Components].AxialGridSize, [Dbf Components].RadialGridSize, [Dbf Components].[Exam Instructions], [Dbf Components].[Accessible Online], [Dbf Components].[Inspection Priority], [Dbf Components].[Inspection LocationID], [Dbf Components].[Initial Inspection Number], [Dbf Components].[Installation Spec], [Dbf Components].[Inspection Interval], [Dbf Components].[Date Install], [Dbf Components].[Date Replaced], [Dbf Components].[Replacement Comments], [Dbf Components].[Replacement WorkOrder], [Dbf Components].PrevComponentID, [Dbf Components].[Reinspection Interval], [Dbf Components].ReinstallationOutage, [Dbf Components].NextScheduledInspection, [Dbf Components].DefaultInspectionMeth, [Dbf Components].[Construction Code], [Dbf Components].[Repair-Replace], [Dbf Components].[Repair-Replace Date1], [Dbf Components].[Repair-Replace WO1], [Dbf Components].[Repair-Replace Date2], [Dbf Components].[Repair-Replace WO2], [Dbf Components].[Repair-Replace Date3], [Dbf Components].[Repair-Replace WO3], [Dbf Components].[Repair-Replace Date4], [Dbf Components].[Repair-Replace WO4], [Dbf Components].[Repair-Replace Date5], [Dbf Components].[Repair-Replace WO5], " +
                                                   "[Dbf Components].[Repair-Replace Date6], [Dbf Components].[Repair-Replace WO6], [Dbf Components].Custom1, [Dbf Components].Custom2, [Dbf Components].Custom3, [Dbf Components].Custom4, [Dbf Components].Custom5, [Dbf Components].Custom6, [Dbf Components].Custom7, [Dbf Components].Custom8, [Dbf Components].Custom9, [Dbf Components].Custom10, [Dbf Components].Custom11, [Dbf Components].Custom12, [Dbf Components].Custom13, [Dbf Components].Custom14, [Dbf Components].Custom15, [Dbf Components].Custom16, [Dbf Components].Custom17, [Dbf Components].Custom18, [Dbf Components].Custom19, [Dbf Components].Custom20, [Dbf Components].[Geometry Code], [Dbf Components].[DNST Geo Code], [Dbf Components].Scheduled, [Dbf Components].Usage, [Dbf Components].Material, [Dbf Components].[Material Det], [Dbf Components].PlantExp, [Dbf Components].DirectDeposit, [Dbf Components].TotBioF, [Dbf Components].ChemTreatF, [Dbf Components].LaSKIndex, [Dbf Components].ImpactLeakF, [Dbf Components].ConType, [Dbf Components].SafetyConseq, [Dbf Components].SizeF, [Dbf Components].FlowRateFactor, [Dbf Components].UseFactor, [Dbf Components].MaterialF, [Dbf Components].PipeOrientF, [Dbf Components].TotDesignF, [Dbf Components].PE, [Dbf Components].DirectDepositF, [Dbf Components].IndirectDepositF, [Dbf Components].CompZF, [Dbf Components].CompZFWOPE, [Dbf Components].ProdConseq, " +
                                                   "[Dbf Components].SafetyConsFactor, [Dbf Components].TotConsq, [Dbf Components].Risk, [Dbf Components].RiskWOPE, [Dbf Components].CalculatePlantExt, [Dbf Components].[Deposit Accumulation], [Dbf Components].[Min Thickness Found], [Dbf Components].[Tmeas From], [Dbf Components].BeginPoint, [Dbf Components].EndPoint, [Dbf Components].MapFeatureID FROM ((([Dbf Components] INNER JOIN [Dbf Zone] ON [Dbf Components].[Zone ID] = [Dbf Zone].[Zone ID]) INNER JOIN [Dbf Lines] ON [Dbf Zone].LineID = [Dbf Lines].LineID) INNER JOIN [Dbf Systems] ON [Dbf Lines].SystemID = [Dbf Systems].SystemID) LEFT JOIN [Dbf Subsystems] ON [Dbf Zone].SubSystemID = [Dbf Subsystems].SubSystemID;";
                        break;
                    }

                case (LayerSystemUse.suSWFlowSegments):
                    {
                        layerdef.BaseFilterQuery = "SELECT [Dbf Zone].[Zone ID], [Dbf Zone].[Flow Segment], [Dbf Zone].[Flow Segment Ref #], [Dbf Zone].[Flow Segment Type], [Dbf Zone].[Flow Segment Description], [Dbf Systems].System, [Dbf Lines].Line, [Dbf Zone].Node, [Dbf Zone].[Flow (gpm)], [Dbf Zone].[Flow Density], [Dbf Zone].[Nominal Pipe Size], [Dbf Zone].[Pipe OD], [Dbf Zone].[Pipe ID], [Dbf Zone].Area, [Dbf Zone].Velocity, [Dbf Zone].Comments, [Dbf Zone].Scheduled, [Dbf Zone].PID, [Dbf Zone].Iso, [Dbf Zone].PRA, [Dbf Zone].LCO, [Dbf Zone].Accessablility, [Dbf Zone].LineSize, [Dbf Zone].Building, [Dbf Zone].Elevation, [Dbf Zone].Row, [Dbf Zone].Col, [Dbf Zone].SR, [Dbf Zone].Schedule, [Dbf Zone].Category, [Dbf Zone].ChemTF, [Dbf Zone].Online, [Dbf Zone].SRB, [Dbf Zone].TAB, [Dbf Zone].APB, [Dbf Zone].IRB, [Dbf Zone].Cl, [Dbf Zone].SO3, [Dbf Zone].HCO2, [Dbf Zone].CO2, [Dbf Zone].Comments2, [Dbf Zone].InspLocation, [Dbf Zone].Drawing1, [Dbf Zone].Drawing2, [Dbf Zone].Unit, [Dbf Zone].[Tnom Main], [Dbf Zone].BeginPoint, [Dbf Zone].EndPoint, [Dbf Zone].MapFeatureID FROM ([Dbf Zone] INNER JOIN [Dbf Lines] ON [Dbf Zone].LineID = [Dbf Lines].LineID) INNER JOIN [Dbf Systems] ON [Dbf Lines].SystemID = [Dbf Systems].SystemID";
                        break;
                    }

                case (LayerSystemUse.suSWLines):
                    {
                        layerdef.BaseFilterQuery = "SELECT [Dbf Lines].LineID, [Dbf Lines].Line, [Dbf Lines].LineNumber, [Dbf Lines].LineDescription, [Dbf Lines].LineClass, [Dbf Systems].System, [Dbf Lines].LineGroup, [Dbf Lines].LinePhase, [Dbf Lines].LineSafetyGrade, [Dbf Lines].LineSize, [Dbf Lines].BoreClass, [Dbf Lines].Category, [Dbf Lines].Criteria, [Dbf Lines].BasisComments, [Dbf Lines].CheckWorksLineName, [Dbf Lines].CheckWorksLineNumber, [Dbf Lines].FailureConsequence, [Dbf Lines].[Level of Susceptability], [Dbf Lines].Temperature, [Dbf Lines].[Steam Quality], [Dbf Lines].[Operating Time], [Dbf Lines].EvalComments, [Dbf Lines].SNMComments, [Dbf Lines].References, [Dbf Lines].PlantExperience, [Dbf Lines].PlantExperienceDescription, [Dbf Lines].IndustryExperience, [Dbf Lines].IndustryExperienceDescription, [Dbf Lines].Comments, [Dbf Lines].Rev, [Dbf Lines].RefNo, [Dbf Lines].[Drawing No], [Dbf Lines].Unit, [Dbf Lines].Usage, [Dbf Lines].BeginPoint, [Dbf Lines].EndPoint, [Dbf Lines].MapFeatureID FROM [Dbf Lines] INNER JOIN [Dbf Systems] ON [Dbf Lines].SystemID = [Dbf Systems].SystemID";
                        break;
                    }
                    
            }
            _layers.Add(layerdef);
            return _layers.Count;
        }


        //----------------------------------
        // Open the Layer Settings Dialog
        //----------------------------------
        public string GetLayerTypeName(LayerSystemUse systemUse)
        {
            if (((int) systemUse >= 1) && ((int) systemUse <= _layernametypescount))
                return _layernametypes[(int) systemUse];
           return "";
        }



        //----------------------------------
        // Open the Layer Settings Dialog
        //----------------------------------
        public void OpenLayerSettings(LayerSystemUse systemUse, ArcEngineClass GISData, ref bool displayChanged)
        {
            // Add the base query if it's a system layer 
            switch (systemUse)
            {
                case (LayerSystemUse.suBPLines):
                    {                        
                        break;
                    }
                case (LayerSystemUse.suBPLineSections):
                    {
                        // First get the layer
                        LayerDef LyrDef = new LayerDef();
                        if (GetLayerDef(systemUse, ref LyrDef))
                        {
                            BPLineSectionLayerSettings lyrform = new BPLineSectionLayerSettings(LyrDef, GISData);
                            if (lyrform.ShowDialog() == DialogResult.OK)
                            {
                                displayChanged = lyrform.DisplayModified;
                                if (displayChanged)
                                {
                                    //if (lyrform.Displaymode == 1)
                                    //{
                                    //    LyrDef.m_SimpleRenderer = lyrform.SimpleRenderer;
                                    //    GISData.SetLayerRenderer(LyrDef);                                       
                                    //}
                                    //else
                                    //{
                                    //    LyrDef.m_classBreaksRenderer = lyrform.LineSectionRender;
                                    //    GISData.SetLayerRenderer(LyrDef);                                        
                                    //}
                                }                                
                            }
                        }
                        break;
                    }
                case (LayerSystemUse.suBPBoringLocations):
                    {
                        break;
                    }
                case (LayerSystemUse.suBPCorrosionProbes):
                    {
                        break;
                    }
            }
        }

    }

    //------------------------------------------------------------------------------
    //------------------------------------------------------------------------------
    public struct GISDBSettings
    {
        //public string GISDBPath;
        public string GISRootFolder;
        public string GISmxdfile;
        public bool GISEnabled;
        public string CWGraphingEXE;
        public bool  SampleSelectionByGroundwater;
        public double MinNorthing;
        public double MaxNorthing;
        public double MinEasting;
        public double MaxEasting;
    }


    //------------------------------------------------------------------------------
    //------------------------------------------------------------------------------
    // Display/Value ListBox Definitions
    // Int type
    class ListBoxNameValueIntObject
    {
        public string ObjectName { get; set; }
        public int ObjectValue { get; set; }
    }

    // string type
    class ListBoxNameValueStringObject
    {
        public string ObjectName { get; set; }
        public string ObjectValue { get; set; }
    }

    // SQLFilter type
    class ListBoxNameValueFilterObject
    {
        public string ObjectName { get; set; }
        public SQLFilterClass ObjectValue { get; set; }
    }


    //------------------------------------------------------------------------------
    // SQL Condition Definition
    //------------------------------------------------------------------------------
    public class SQLCondition
    {
        public string logicaloperator;
        public string fieldname;
        public string condition;
        public string value;
        public string fieldtype;

        public SQLCondition()
        {
            logicaloperator = "";
            fieldname = "";
            condition = "";
            value = "";
            fieldtype = "";
        }
    }

    //------------------------------------------------------------------------------
    // SQL Filter Class
    //------------------------------------------------------------------------------
    public class SQLFilterClass
    {
        // Private Properties
        private List<SQLCondition> _conditions;

        // Public Properties
        public int ID = 0;
        public LayerSystemUse LayerIndex = LayerSystemUse.suBPLines;
        public string FilterName = "";
        public string FilterDescription = "";
        public int ConditionCount { get { return _conditions.Count; } }

        //-----------------------------
        // Constructor
        // Overload 1
        //-----------------------------
        public SQLFilterClass()
        {
            _conditions = new List<SQLCondition>();
        }

        //-----------------------------
        // Constructor - copies an existing filter
        // Overload 2
        //-----------------------------
        public SQLFilterClass(SQLFilterClass filtertocopy)
        {
            _conditions = new List<SQLCondition>();
            if (filtertocopy != null)
            {
                Assign(filtertocopy);
            }
        }

        //-----------------------------
        // Assigns a filter to an existing filter
        //-----------------------------
        public bool Assign(SQLFilterClass filtertocopy)
        {
            _conditions.Clear();
            if (filtertocopy != null)
            {
                ID = filtertocopy.ID;
                LayerIndex = filtertocopy.LayerIndex;
                FilterName = filtertocopy.FilterName;
                FilterDescription = filtertocopy.FilterDescription;
                int i;
                for (i = 0; i <= (filtertocopy.ConditionCount - 1); i++)
                {
                    AddCondition(filtertocopy._conditions[i].fieldname, filtertocopy._conditions[i].condition, filtertocopy._conditions[i].value, filtertocopy._conditions[i].fieldtype, filtertocopy._conditions[i].logicaloperator);
                }
                return true;
            }
            else
                return false;
        }


        //-----------------------------
        // Adds a new Condition
        // Overload 1
        //-----------------------------
        public bool AddCondition(string fieldname, string condition, string value, string fieldtype, string logicaloperator)
        {
            if ((fieldname.Length == 0) || (condition.Length == 0) || (value.Length == 0)) return false;
            SQLCondition cond = new SQLCondition();
            cond.condition = condition;
            cond.fieldname = fieldname;
            cond.value = value;
            cond.fieldtype = fieldtype;
            cond.logicaloperator = logicaloperator;
            _conditions.Add(cond);
            return true;
        }

        //-----------------------------
        // Adds a new Condition
        // Overload 2
        //-----------------------------
        public bool AddCondition(SQLCondition cond)
        {
            if ((cond.fieldname.Length == 0) || (cond.condition.Length == 0) || (cond.value.Length == 0)) return false;
            SQLCondition newcond = new SQLCondition();
            newcond.condition = cond.condition;
            newcond.fieldname = cond.fieldname;
            newcond.value = cond.value;
            newcond.fieldtype = cond.fieldtype;
            newcond.logicaloperator = cond.logicaloperator;
            _conditions.Add(cond);
            return true;
        }

        //-----------------------------
        // Get a Condition
        //-----------------------------
        public bool GetCondition(int index, ref SQLCondition cond)
        {
            if ((index < 0) || (index > (_conditions.Count - 1))) return false;
            cond = _conditions[index];
            return true;
        }

        //-----------------------------
        // Clears the 
        //-----------------------------
        public void Clear()
        {

        }
    }


    //------------------------------------------------------------------------------
    // Feature selection set metadata class
    //------------------------------------------------------------------------------
    public class SelectedItemData
    {
        public int AttribObjectID;
        public string CWKeyValue;
        public string layername;
        public int nodelevel;
        public string featuredescription;
    }


    //------------------------------------------------------------------------------
    // Root Node selection set metadata class
    //------------------------------------------------------------------------------
    public class SelectedRootNodeData
    {
        public string layername;
        public LayerSystemUse layerindex;
        public SQLFilterClass activefilter = null;
    }


    //------------------------------------------------------------------------------
    // Risk Ranking Contructs
    //------------------------------------------------------------------------------
    public struct RiskRankingValue
    {
        public string RRDesc;
        public int RRValue;
    }


    //------------------------------------------------------------------------------
    // Risk Ranking Range Node
    //------------------------------------------------------------------------------
    public class RiskRankingRange
    {
        public RiskRankingRange(int lower, int upper, byte rcolor, byte gcolor, byte bcolor)
        {
            Lower = lower;
            Upper = upper;
            rColor = rcolor;
            gColor = gcolor;
            bColor = bcolor;
        }
        public int Lower;
        public int Upper;
        public byte rColor;
        public byte gColor;
        public byte bColor;
    }


} // Namespace
