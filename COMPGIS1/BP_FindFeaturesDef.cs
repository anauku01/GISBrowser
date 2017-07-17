using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace COMPGIS1
{
    class FindFeatureClass
    {
        public List<string> Items = new List<string>();
        private string _featuredesc = "";
        public string Description { get { return _featuredesc; } set { } }

        public FindFeatureClass(string description)
        {
            _featuredesc = description;
        }
    }



    class BP_FindFeaturesDef
    {
        
        private List<FindFeatureClass> _features = new List<FindFeatureClass>();


        //......................................
        // Constructor
        //......................................
        public BP_FindFeaturesDef()
        {
            LoadValues();
        }

        //......................................
        // Load the Features in to the array
        //......................................
        private void LoadValues()
        {
            _features.Clear();
            FindFeatureClass fc = new FindFeatureClass("Lines");
            fc.Items.Add("Lines containing ");

        }
    }
}
