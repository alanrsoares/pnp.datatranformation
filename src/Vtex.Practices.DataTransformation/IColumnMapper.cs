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
        IColumnMapper<T> Map(int? index, string propertyName, string headerText, CellType cellType, Func<object, object> customTranformationAction);
        IColumnMapper<T> Map(string propertyName, string headerText, CellType cellType, Func<object, object> customTranformationAction);
        IColumnMapper<T> Map(int index, Func<object, object> customTranformationAction);
        IColumnMapper<T> Map(int index, string headerText);
        IColumnMapper<T> Map(string propertyName, Func<object, object> customTranformationAction);
        IColumnMapper<T> Map(string propertyName, string headerText, CellType cellType);
        IColumnMapper<T> Map(string propertyName, string headerText);
        IColumnMapper<T> Map(string propertyName, CellType cellType);
        IColumnMapper<T> Map(string propertyName);
        IColumnMapper<T> Unmap(string propertyName);
        IColumnMapper<T> Unmap(int index);
    }
}