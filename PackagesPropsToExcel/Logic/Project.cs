using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace Logic
{
    [Serializable]
    [XmlRoot("Project", Namespace = "http://schemas.microsoft.com/developer/msbuild/2003")]
    public class Project
    {
        [XmlElement("ItemGroup")]
        public ItemGroup ItemGroup { get; set; }
    }

    [Serializable]
    public class ItemGroup
    {
        [XmlElement("PackageReference")]
        public List<PackageReference> PackageReferences { get; set; }
    }

     [Serializable]
    public class PackageReference
    {
        public string Update { get; set; }
        public string Version { get; set; }
    }
}