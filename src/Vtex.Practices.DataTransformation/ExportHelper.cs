using System.Collections.Generic;

namespace Vtex.Practices.DataTransformation
{
    public static class ExportHelper
    {
        public static IFormatData Export<T>(IEnumerable<T> data) where T : new()
        {
            var mapper = ColumnMapper<T>.Factory.CreateNew();
            mapper.AutoMapColumns();
            return new DataFormatter<T>(mapper, data);
        }

        public static IFormatData Export<T>(IEnumerable<T> data, IColumnMapper<T> mapper) where T : new()
        {
            return new DataFormatter<T>(mapper, data);
        }
    }
}