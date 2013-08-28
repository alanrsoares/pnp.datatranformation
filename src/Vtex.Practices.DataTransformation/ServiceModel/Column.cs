using System;
using NPOI.SS.UserModel;

namespace Vtex.Practices.DataTransformation.ServiceModel
{
    public class Column
    {
        public int Index { get; set; }
        public string PropertyName { get; set; }
        public Type Type { get; set; }
        public CellType? CellType { get; set; }
        public string HeaderText { get; set; }
        public Func<object, object> CustomTransformAction { get; set; }

        public Type UnderLyingType 
        {
            get { return Nullable.GetUnderlyingType(this.Type); }
        }

        public bool IsNullable 
        {
            get { return UnderLyingType != null; }
        }
    }
}