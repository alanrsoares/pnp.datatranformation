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
        private readonly PropertyInfo[] _properties;

        public List<Column> Columns { get; private set; }

        public ColumnMapper()
        {
            _properties = typeof(T).GetProperties();

            Columns = new List<Column>();
        }

        public IColumnMapper<T> MapColumn(Column column)
        {
            if (string.IsNullOrWhiteSpace(column.HeaderText))
                column.HeaderText = column.PropertyName;

            MapColumn(column.Index, column.PropertyName, column.HeaderText, column.CellType.GetValueOrDefault(CellType.Unknown), column.CustomTransformAction);

            return this;
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
        public IColumnMapper<T> MapColumn(int index, string propertyName, string headerText, CellType cellType, Func<object, object> customTranformationAction)
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

        public IColumnMapper<T> MapColumn(string propertyName, string headerText, CellType cellType, Func<object, object> customTranformationAction)
        {
            return MapColumn(GetColumnIndex(propertyName), propertyName, headerText, cellType, customTranformationAction);
        }

        public IColumnMapper<T> MapColumn(string propertyName, Func<object, object> customTranformationAction)
        {
            MapColumn(new Column
                {
                    Index = GetColumnIndex(propertyName),
                    CustomTransformAction = customTranformationAction
                });
            return this;
        }

        private int GetColumnIndex(string propertyName)
        {
            var existingColumn = Columns.FirstOrDefault(c => c.PropertyName.Equals(propertyName));

            return existingColumn == null ? Columns.Count : existingColumn.Index;
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

        private static Type GetType(string propertyName)
        {
            return typeof(T).GetProperty(propertyName).PropertyType;
        }

        public IColumnMapper<T> AutoMapColumns()
        {
            var index = -1;

            Columns.AddRange(_properties.Select(property =>
            {
                index++;

                return new Column
                    {
                        Index = index,
                        PropertyName = property.Name,
                        HeaderText = property.Name,
                        CellType = GetCellType(property.PropertyType),
                        Type = property.PropertyType
                    };
            }).ToList());

            return this;
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
    }
}