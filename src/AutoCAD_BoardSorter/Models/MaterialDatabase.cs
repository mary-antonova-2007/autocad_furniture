using System.Collections.Generic;
using System.Runtime.Serialization;

namespace AutoCAD_BoardSorter.Models
{
    [DataContract]
    internal sealed class MaterialDatabase
    {
        [DataMember(Order = 1)]
        public List<BoardMaterialData> BoardMaterials { get; set; } = new List<BoardMaterialData>();

        [DataMember(Order = 2)]
        public List<CoatingMaterialData> CoatingMaterials { get; set; } = new List<CoatingMaterialData>();

        [DataMember(Order = 3)]
        public List<MaterialCategoryData> BoardCategories { get; set; } = new List<MaterialCategoryData>();

        [DataMember(Order = 4)]
        public List<MaterialCategoryData> CoatingCategories { get; set; } = new List<MaterialCategoryData>();

        [DataMember(Order = 5)]
        public MaterialDatabaseUiLayout UiLayout { get; set; } = new MaterialDatabaseUiLayout();
    }

    [DataContract]
    internal sealed class BoardMaterialData
    {
        [DataMember(Order = 1)] public string Code { get; set; }
        [DataMember(Order = 2)] public string Name { get; set; }
        [DataMember(Order = 3)] public string CalculationType { get; set; }
        [DataMember(Order = 4)] public string CategoryCode { get; set; }
        [DataMember(Order = 5)] public List<MaterialFormatData> Formats { get; set; } = new List<MaterialFormatData>();
        [DataMember(Order = 6)] public string DefaultVisibleEdgeCoatingCode { get; set; }
        [DataMember(Order = 7)] public string DefaultHiddenEdgeCoatingCode { get; set; }
        [DataMember(Order = 8)] public string FrontFaceCoatingCode { get; set; }
        [DataMember(Order = 9)] public string BackFaceCoatingCode { get; set; }
        [DataMember(Order = 10)] public string Note { get; set; }

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(Name) ? Code : Name;
        }
    }

    [DataContract]
    internal sealed class MaterialFormatData
    {
        [DataMember(Order = 1)] public string Code { get; set; }
        [DataMember(Order = 2)] public string FormatType { get; set; }
        [DataMember(Order = 3)] public double Length { get; set; }
        [DataMember(Order = 4)] public double Width { get; set; }
        [DataMember(Order = 5)] public double Thickness { get; set; }
        [DataMember(Order = 6)] public string Note { get; set; }

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(Code) ? string.Format("{0:0.###} x {1:0.###} x {2:0.###}", Length, Width, Thickness) : Code;
        }
    }

    [DataContract]
    internal sealed class CoatingMaterialData
    {
        [DataMember(Order = 1)] public string Code { get; set; }
        [DataMember(Order = 2)] public string Name { get; set; }
        [DataMember(Order = 3)] public string CalculationType { get; set; }
        [DataMember(Order = 4)] public double Thickness { get; set; }
        [DataMember(Order = 5)] public string CategoryCode { get; set; }
        [DataMember(Order = 6)] public string Note { get; set; }

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(Name) ? Code : Name;
        }
    }

    [DataContract]
    internal sealed class MaterialCategoryData
    {
        [DataMember(Order = 1)] public string Code { get; set; }
        [DataMember(Order = 2)] public string Name { get; set; }
        [DataMember(Order = 3)] public string ParentCode { get; set; }
        [DataMember(Order = 4)] public string Note { get; set; }

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(Name) ? Code : Name;
        }
    }

    [DataContract]
    internal sealed class MaterialDatabaseUiLayout
    {
        [DataMember(Order = 1)] public double BoardSidebarWidth { get; set; } = 280.0;
        [DataMember(Order = 2)] public double BoardTreeHeight { get; set; } = 260.0;
        [DataMember(Order = 3)] public double CoatingSidebarWidth { get; set; } = 280.0;
        [DataMember(Order = 4)] public double CoatingTreeHeight { get; set; } = 260.0;
    }
}
