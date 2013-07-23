using System;
using NPOI.SS.UserModel;

namespace Vtex.Practices.DataTransformation.xls
{
    public interface IColumnMapper<T>
    {
        IColumnMapper<T> AutoMapColumns();
        IColumnMapper<T> MapColumn(Column column);
        IColumnMapper<T> MapColumn(int index, string propertyName, string headerText, CellType cellType, Func<object, object> customTranformationAction);
        IColumnMapper<T> MapColumn(string propertyName, string headerText, CellType cellType, Func<object, object> customTranformationAction);
        IColumnMapper<T> MapColumn(string propertyName, Func<object, object> customTranformationAction);
        IColumnMapper<T> MapColumn(string propertyName, string headerText, CellType cellType);
        IColumnMapper<T> MapColumn(string propertyName, string headerText);
        IColumnMapper<T> MapColumn(string propertyName, CellType cellType);
        IColumnMapper<T> MapColumn(string propertyName);
    }
}