using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using NPOI.SS.UserModel;
using Vtex.Practices.DataTransformation.ServiceModel;

namespace Vtex.Practices.DataTransformation
{
    public interface IColumnMapper<T> where T : new()
    {
        List<Column> Columns { get; }

        IEnumerable<PropertyInfo> Properties { get; }

        DataHandler<T> DataHandler { get; }

        IColumnMapper<T> AutoMapColumns();

        IColumnMapper<T> Map(Column column);

        IColumnMapper<T> Map(String propertyName,
                             String headerText,
                             bool isReadOnly = false,
                             CellType cellType = CellType.Unknown,
                             Func<object, object> customEncoder = null,
                             Func<object, object> customDecoder = null);

        IColumnMapper<T> Map<TProperty>(Expression<Func<T, TProperty>> property,
                                        String headerText,
                                        bool isReadOnly = false,
                                        CellType cellType = CellType.Unknown,
                                        Func<object, object> customEncoder = null,
                                        Func<object, object> customDecoder = null);

        IColumnMapper<T> Map(String propertyName,
                             Func<object, object> customEncoder = null,
                             Func<object, object> customDecoder = null);

        IColumnMapper<T> Map(String propertyName, CellType cellType);

        IColumnMapper<T> Map<TProperty>(Expression<Func<T, TProperty>> property, CellType cellType);

        IColumnMapper<T> Map(IEnumerable<String> propertyNames);

        IColumnMapper<T> Unmap(String propertyName);

        IColumnMapper<T> Unmap(IEnumerable<String> propertyNames);

        IColumnMapper<T> Map<TProperty>(Expression<Func<T, TProperty>> property);
    }
}