using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using Vtex.Practices.DataTransformation.ServiceModel;

namespace Vtex.Practices.DataTransformation
{
    public interface IColumnMapper<T> where T : new()
    {
        List<Column> Columns { get; }

        DataHandler<T> DataHandler { get; }

        IColumnMapper<T> AutoMapColumns();

        IColumnMapper<T> Map(Column column);

        IColumnMapper<T> Map(string propertyName,
            string headerText,
            CellType cellType = CellType.Unknown,
            Func<object, object> customEncoder = null,
            Func<object, object> customDecoder = null);

        IColumnMapper<T> Map(string propertyName,
            Func<object, object> customEncoder = null,
            Func<object, object> customDecoder = null);

        IColumnMapper<T> Map(string propertyName, CellType cellType);

        IColumnMapper<T> Map(IEnumerable<string> propertyNames);

        IColumnMapper<T> Unmap(string propertyName);
    }
}