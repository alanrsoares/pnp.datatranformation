using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using NPOI.SS.UserModel;
using Vtex.Practices.DataTransformation.Exceptions;
using Vtex.Practices.DataTransformation.ServiceModel;

namespace Vtex.Practices.DataTransformation
{
    public class ColumnMapper<T> : IColumnMapper<T> where T : new()
    {
        public IEnumerable<PropertyInfo> Properties { get; private set; }

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
            get
            {
                return new ColumnMapperFactory<T>();
            }
        }

        public ColumnMapper()
        {
            Properties = typeof(T).GetProperties().ToList();

            Columns = new List<Column>();
        }

        public IColumnMapper<T> AutoMapColumns()
        {
            Properties.ToList().ForEach(property => Map(property.Name));

            return this;
        }

        public IColumnMapper<T> Map(Column column)
        {
            Column existingColumn;

            if (Columns.Any() &&
                !string.IsNullOrWhiteSpace(column.PropertyName) &&
                (existingColumn = GetExistingColumn(column.PropertyName)) != null)
            {
                var edited = existingColumn;

                if (!string.IsNullOrWhiteSpace(column.HeaderText))
                    edited.HeaderText = column.HeaderText;

                edited.CellType = column.CellType
                    ?? existingColumn.CellType
                    ?? GetCellType(column.PropertyName);

                edited.Type = column.Type
                    ?? existingColumn.Type
                    ?? GetType(column.PropertyName);

                if (column.CustomEncoder != null)
                    edited.CustomEncoder = column.CustomEncoder;

                if (column.CustomDecoder != null)
                    edited.CustomDecoder = column.CustomDecoder;

                if (column.ListSeparator != null)
                    edited.ListSeparator = column.ListSeparator;

                Columns[existingColumn.Index.GetValueOrDefault()] = edited;
            }
            else
            {
                if (!column.Index.HasValue)
                    column.Index = Columns.Count;

                if (string.IsNullOrWhiteSpace(column.HeaderText))
                    column.HeaderText = column.PropertyName;

                if (column.CellType == null || column.CellType == CellType.Unknown)
                    column.CellType = GetCellType(column.PropertyName);

                if (column.Type == null)
                    column.Type = GetType(column.PropertyName);

                if (column.ListSeparator == null)
                    column.ListSeparator = ";";

                Columns.Add(column);
            }

            return this;
        }

        public IColumnMapper<T> Map(string propertyName, string headerText, bool isReadOnly = false, CellType cellType = CellType.Unknown, 
            Func<object, object> customEncoder = null, Func<object, object> customDecoder = null)
        {
            var column = new Column
            {
                PropertyName = propertyName,
                HeaderText = headerText,
                CellType = cellType,
                CustomEncoder = customEncoder,
                CustomDecoder = customDecoder,
                IsReadOnly = isReadOnly
            };
            return Map(column);
        }

        public IColumnMapper<T> Map<TProperty>(Expression<Func<T, TProperty>> property, string headerText, bool isReadOnly = false, CellType cellType = CellType.Unknown,
            Func<object, object> customEncoder = null, Func<object, object> customDecoder = null)
        {
            var propertyInfo = GetPropertyInfo(property);

            return Map(propertyInfo.Name, headerText, isReadOnly, cellType, customEncoder, customDecoder);
        }

        public IColumnMapper<T> Map(string propertyName,
                                    CellType cellType)
        {
            return Map(propertyName, propertyName, false, cellType);
        }

        public IColumnMapper<T> Map<TProperty>(Expression<Func<T, TProperty>> property, CellType cellType)
        {
            var propertyInfo = GetPropertyInfo(property);
            return Map(propertyInfo.Name, cellType);
        }

        public IColumnMapper<T> Map(string propertyName,
                                    Func<object, object> customEncoder = null,
                                    Func<object, object> customDecoder = null)
        {
            return Map(propertyName, propertyName, false, CellType.Unknown, customEncoder, customDecoder);
        }

        public IColumnMapper<T> Map(IEnumerable<string> propertyNames)
        {
            propertyNames.ToList().ForEach(propertyName => Map(propertyName));
            return this;
        }

        public IColumnMapper<T> Map<TProperty>(Expression<Func<T, TProperty>> property)
        {
            var propertyInfo = GetPropertyInfo(property);
            return Map(propertyInfo.Name);
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

            return ReArrangeColumns();
        }

        public IColumnMapper<T> Unmap(IEnumerable<string> propertyNames)
        {
            foreach (var propertyName in propertyNames)
            {
                Unmap(propertyName);
            }

            return ReArrangeColumns();
        }

        private IColumnMapper<T> ReArrangeColumns()
        {
            Columns = Columns.Select(column =>
            {
                column.Index = Columns.IndexOf(column);
                return column;
            }).ToList();
            return this;
        }

        private static Type GetType(string propertyName)
        {
            var property = typeof(T).GetProperty(propertyName);

            if (property == null)
                throw new InvalidPropertyException(string.Format("Structure {0} does not have given property: {1}", typeof(T).Name, propertyName));

            return property.PropertyType;
        }

        private CellType GetCellType(string propertyName)
        {
            var property = Properties.FirstOrDefault(p => p.Name.Equals(propertyName));

            if (property != null)
                return GetCellType(property.PropertyType);

            var message = string.Format("Invalid PropertyName: {0} does not correspond to any dto property", propertyName);

            throw new InvalidPropertyException(message);
        }

        private CellType GetCellType(Type propertyType)
        {
            var propertyTypeName = propertyType.Name;

            var underLyingType = Nullable.GetUnderlyingType(propertyType);

            if (underLyingType != null)
            {
                propertyTypeName = underLyingType.Name;
            }

            if (!propertyType.Name.Equals("Nullable`1") && (propertyType.IsGenericType || propertyType.IsArray))
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
                case "Boolean":
                    return CellType.BOOLEAN;
                default:
                    return CellType.Unknown;
            }
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

        private static PropertyInfo GetPropertyInfo<TSource, TProperty>(Expression<Func<TSource, TProperty>> propertyLambda)
        {
            var type = typeof(TSource);

            var member = propertyLambda.Body as MemberExpression;

            if (member == null)
                throw new ArgumentException(string.Format("Expression '{0}' refers to a method, not a property.", propertyLambda));

            var propInfo = member.Member as PropertyInfo;

            if (propInfo == null)
                throw new ArgumentException(string.Format("Expression '{0}' refers to a field, not a property.", propertyLambda));

            if (type != propInfo.ReflectedType && !type.IsSubclassOf(propInfo.ReflectedType))
                throw new ArgumentException(string.Format("Expresion '{0}' refers to a property that is not from type {1}.", propertyLambda, type));

            return propInfo;
        }
    }
}