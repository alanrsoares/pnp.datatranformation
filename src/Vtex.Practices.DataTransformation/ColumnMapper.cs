using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using NPOI.SS.UserModel;
using Vtex.Practices.DataTransformation.xls;

namespace Vtex.Practices.DataTransformation
{
    public class ColumnMapper<T> : IColumnMapper<T>
    {
        private readonly List<PropertyInfo> _properties;

        public List<Column> Columns { get; private set; }

        public static ColumnMapperFactory<T> Factory
        {
            get { return new ColumnMapperFactory<T>(); }
        }

        public ColumnMapper()
        {
            _properties = typeof(T).GetProperties().ToList();

            Columns = new List<Column>();
        }

        public IColumnMapper<T> MapColumn(Column column)
        {
            if (string.IsNullOrWhiteSpace(column.HeaderText))
                column.HeaderText = column.PropertyName;

            return MapColumn(column.Index,
                             column.PropertyName,
                             column.HeaderText,
                             column.CellType.GetValueOrDefault(CellType.Unknown),
                             column.CustomTransformAction);
        }

        /// <summary>
        /// Will add or overwrite column at given index
        /// </summary>
        /// <param name="index"></param>
        /// <param name="propertyName"></param>
        /// <param name="headerText"></param>
        /// <param name="cellType"></param>
        /// <param name="customTranformationAction"></param>
        /// <returns></returns>
        public IColumnMapper<T> MapColumn(int index, string propertyName, string headerText, CellType cellType,
                                          Func<object, object> customTranformationAction)
        {
            var newColumn = new Column
                {
                    Index = index,
                    PropertyName = propertyName,
                    HeaderText = headerText,
                    CellType = cellType,
                    Type = GetType(propertyName),
                    CustomTransformAction = customTranformationAction
                };

            if (Columns.Count > index && Columns[index] != null)
                Columns[index] = newColumn;
            else
                Columns.Add(newColumn);

            return this;
        }

        public IColumnMapper<T> MapColumn(string propertyName, string headerText, CellType cellType,
                                          Func<object, object> customTranformationAction)
        {
            return MapColumn(GetColumnIndex(propertyName), propertyName, headerText, cellType, customTranformationAction);
        }

        public IColumnMapper<T> MapColumn(int index, Func<object, object> customTranformationAction)
        {
            var existingColumn = Columns[index];

            return MapColumn(existingColumn.Index,
                             existingColumn.PropertyName,
                             existingColumn.HeaderText,
                             existingColumn.CellType.GetValueOrDefault(CellType.Unknown),
                             customTranformationAction);
        }

        public IColumnMapper<T> MapColumn(string propertyName, Func<object, object> customTranformationAction)
        {
            return MapColumn(GetColumnIndex(propertyName), propertyName, propertyName, GetCellType(propertyName),
                             customTranformationAction);
        }

        public IColumnMapper<T> MapColumn(string propertyName, string headerText, CellType cellType)
        {
            return MapColumn(propertyName, headerText, cellType, null);
        }

        public IColumnMapper<T> MapColumn(string propertyName, string headerText)
        {
            return MapColumn(propertyName, headerText, CellType.Unknown);
        }

        public IColumnMapper<T> MapColumn(string propertyName, CellType cellType)
        {
            return MapColumn(propertyName, propertyName, cellType);
        }

        public IColumnMapper<T> MapColumn(string propertyName)
        {
            return MapColumn(propertyName, CellType.Unknown);
        }

        public IColumnMapper<T> AutoMapColumns()
        {
            _properties.ForEach(property =>
                                MapColumn(property.Name, GetCellType(property.PropertyType)));

            return this;
        }

        private static Type GetType(string propertyName)
        {
            return typeof(T).GetProperty(propertyName).PropertyType;
        }

        private int GetColumnIndex(string propertyName)
        {
            var existingColumn = Columns.FirstOrDefault(c => c.PropertyName.Equals(propertyName));

            return existingColumn == null ? Columns.Count : existingColumn.Index;
        }

        private CellType GetCellType(Type propertyType)
        {
            switch (propertyType.Name)
            {
                case "String":
                    return CellType.STRING;
                case "Int32":
                case "Double":
                case "Decimal":
                    return CellType.NUMERIC;
                default:
                    return CellType.Unknown;
            }
        }

        private CellType GetCellType(string propertyName)
        {
            var property = _properties.FirstOrDefault(p => p.Name == propertyName);
            if (property == null)
                throw new IndexOutOfRangeException(
                    string.Format("Invalid PropertyName: {0} does not correspond to any dto property", propertyName));

            return GetCellType(property.PropertyType);
        }
    }

    public class ColumnMapperFactory<T>
    {
        public ColumnMapper<T> CreateNew()
        {
            return new ColumnMapper<T>();
        }
    }
}