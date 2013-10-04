using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using NPOI.SS.UserModel;
using Vtex.Practices.DataTransformation.Exceptions;
using Vtex.Practices.DataTransformation.ServiceModel;

namespace Vtex.Practices.DataTransformation
{
    public class ColumnMapper<T> : IColumnMapper<T> where T : new()
    {
        private readonly List<PropertyInfo> _properties;

        public List<Column> Columns { get; private set; }

        public DataHandler<T> DataHandler
        {
            get
            {
                return DataHandlerFactory.NewFromColumnMapper(this);
            }
        }

        public static ColumnMapperFactory<T> Factory
        {
            get { return new ColumnMapperFactory<T>(); }
        }

        public ColumnMapper()
        {
            _properties = typeof(T).GetProperties().ToList();

            Columns = new List<Column>();
        }

        public IColumnMapper<T> Map(Column column)
        {
            var headerText = column.HeaderText;

            var propertyName = column.PropertyName;

            var index = column.Index;

            var existingColumn = GetExistingColumn(index);

            var newColumn = new Column
            {
                Index = index,
                PropertyName = propertyName,
                HeaderText = string.IsNullOrWhiteSpace(headerText) ? propertyName : headerText,
                CellType = column.CellType ?? GetCellType(propertyName),
                Type = GetType(propertyName),
                CustomEncoder = column.CustomEncoder,
                CustomDecoder = column.CustomDecoder,
                ListSeparator = column.ListSeparator
            };

            if (existingColumn != null)
                Columns[index.GetValueOrDefault()] = newColumn;
            else
                Columns.Add(newColumn);

            return this;
        }

        public IColumnMapper<T> Unmap(int index)
        {
            if (GetExistingColumn(index) != null)
            {
                Columns.RemoveAt(index);
            }
            else
            {
                var message = string.Format("Invalid property index: {0}. No property found at given index.", index);
                throw new IndexOutOfRangeException(message);
            }

            return this;
        }

        public IColumnMapper<T> Unmap(string propertyName)
        {
            var column = GetExistingColumn(propertyName);

            if (column != null)
            {
                Columns.Remove(column);
            }
            else
            {
                var message = string.IsNullOrWhiteSpace(propertyName)
                    ? "Argument \"prpertyName\" cannot have either null or empty value."
                    : string.Format("You are trying to unmap an invalid property: \"{0}\".", propertyName);

                throw new InvalidPropertyException(message);
            }

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
        public IColumnMapper<T> Map(int? index, string propertyName, string headerText, CellType cellType, Func<object, object> customTranformationAction)
        {
            var newColumn = new Column
                {
                    Index = index,
                    PropertyName = propertyName,
                    HeaderText = headerText,
                    CellType = cellType,
                    Type = GetType(propertyName),
                    CustomEncoder = customTranformationAction
                };

            if (index != null && Columns.Count > index && Columns[index.Value] != null)
                Columns[index.Value] = newColumn;
            else
                Columns.Add(newColumn);

            return this;
        }

        public IColumnMapper<T> Map(string propertyName, string headerText, CellType cellType,
                                          Func<object, object> customTranformationAction)
        {
            return Map(GetColumnIndex(propertyName), propertyName, headerText, cellType, customTranformationAction);
        }

        public IColumnMapper<T> Map(int index, Func<object, object> customTranformationAction)
        {
            var existingColumn = Columns[index];

            return Map(existingColumn.Index,
                             existingColumn.PropertyName,
                             existingColumn.HeaderText,
                             existingColumn.CellType.GetValueOrDefault(CellType.Unknown),
                             customTranformationAction);
        }

        public IColumnMapper<T> Map(int index, string headerText)
        {
            var existingColumn = Columns[index];

            return Map(index,
                             existingColumn.PropertyName,
                             headerText,
                             existingColumn.CellType.GetValueOrDefault(CellType.Unknown),
                             existingColumn.CustomEncoder);
        }

        public IColumnMapper<T> Map(string propertyName, Func<object, object> customTranformationAction)
        {
            return Map(GetColumnIndex(propertyName),
                             propertyName, propertyName,
                             GetCellType(propertyName),
                             customTranformationAction);
        }

        public IColumnMapper<T> Map(string propertyName, string headerText, CellType cellType)
        {
            return Map(propertyName, headerText, cellType, null);
        }

        public IColumnMapper<T> Map(string propertyName, string headerText)
        {
            return Map(propertyName, headerText, CellType.Unknown);
        }

        public IColumnMapper<T> Map(string propertyName, CellType cellType)
        {
            return Map(propertyName, propertyName, cellType);
        }

        public IColumnMapper<T> Map(string propertyName)
        {
            return Map(propertyName, GetCellType(propertyName));
        }

        public IColumnMapper<T> AutoMapColumns()
        {
            _properties.ForEach(property =>
                Map(property.Name, GetCellType(property.PropertyType)));

            return this;
        }

        private static Type GetType(string propertyName)
        {
            var property = typeof(T).GetProperty(propertyName);

            if (property == null)
                throw new InvalidPropertyException(string.Format("Structure {0} does not have given property: {1}", typeof(T).Name, propertyName));

            return property.PropertyType;
        }

        private int? GetColumnIndex(string propertyName)
        {
            var existingColumn = Columns.FirstOrDefault(c => c.PropertyName.Equals(propertyName));

            return existingColumn == null ? Columns.Count : existingColumn.Index;
        }

        private static CellType GetCellType(Type propertyType)
        {
            var propertyTypeName = propertyType.Name;

            var underLyingType = Nullable.GetUnderlyingType(propertyType);

            if (underLyingType != null)
            {
                propertyTypeName = underLyingType.Name;
            }

            if (propertyType.IsGenericType || propertyType.IsArray)
            {
                return CellType.STRING;
            }

            switch (propertyTypeName)
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

            if (property != null)
                return GetCellType(property.PropertyType);

            var message = string.Format("Invalid PropertyName: {0} does not correspond to any dto property", propertyName);

            throw new InvalidPropertyException(message);
        }

        private Column GetExistingColumn(int? index)
        {
            Column existingColumn = null;

            if (index != null && Columns.Any() && Columns.Count > index)
            {
                existingColumn = Columns[index.Value];
            }

            return existingColumn;
        }

        private Column GetExistingColumn(string propertyName)
        {
            Column existingColumn = null;

            if (!string.IsNullOrWhiteSpace(propertyName) && Columns.Any())
            {
                existingColumn = Columns.FirstOrDefault(column => column.PropertyName.Equals(propertyName));
            }

            return existingColumn;
        }
    }
}