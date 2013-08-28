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
        IColumnMapper<T> MapColumn(Column column);
        IColumnMapper<T> MapColumn(int index, string propertyName, string headerText, CellType cellType, Func<object, object> customTranformationAction);
        IColumnMapper<T> MapColumn(string propertyName, string headerText, CellType cellType, Func<object, object> customTranformationAction);
        IColumnMapper<T> MapColumn(int index, Func<object, object> customTranformationAction);
        IColumnMapper<T> MapColumn(int index, string headerText);
        IColumnMapper<T> MapColumn(string propertyName, Func<object, object> customTranformationAction);
        IColumnMapper<T> MapColumn(string propertyName, string headerText, CellType cellType);
        IColumnMapper<T> MapColumn(string propertyName, string headerText);
        IColumnMapper<T> MapColumn(string propertyName, CellType cellType);
        IColumnMapper<T> MapColumn(string propertyName);
    }
}