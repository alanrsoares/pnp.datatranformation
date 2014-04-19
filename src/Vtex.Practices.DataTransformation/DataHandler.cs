using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;


namespace Vtex.Practices.DataTransformation
{
    using Column = ServiceModel.Column;

    public class DataHandler<T> where T : new()
    {
        private readonly IColumnMapper<T> _mapper;

        public HSSFWorkbook Workbook { get; private set; }

        public DataHandler(IColumnMapper<T> columnMapper)
        {
            _mapper = columnMapper;
        }

        public HSSFWorkbook EncodeDataToWorkbook(IEnumerable<T> data)
        {
            Workbook = InitializeWorkbook();

            var index = 0;

            const int maxSpreadsheetSize = 65535;

            var dataItems = data as IList<T> ?? data.ToList();

            var spreadSheet = GetPagedSpreadSheetData(dataItems, index, maxSpreadsheetSize);

            while (spreadSheet.Any())
            {
                index++;

                ProcessSpreadSheet(index, spreadSheet);

                spreadSheet = GetPagedSpreadSheetData(dataItems, index, maxSpreadsheetSize);
            }
    
            return Workbook;
        }

        private void ProcessSpreadSheet(int index, IEnumerable<T> spreadSheet)
        {
            var sheet = Workbook.CreateSheet("Sheet" + index);
            
            sheet.ProtectSheet(string.Empty);

            CreateAndPopulateHeader(sheet);

            SetColumnDefaults(sheet);

            CreateAndPopulateRows(sheet, spreadSheet);

            _mapper.Columns.ForEach(c => sheet.AutoSizeColumn(c.Index.GetValueOrDefault()));
        }

        private static List<T> GetPagedSpreadSheetData(IEnumerable<T> dataItems, int index, int maxSpreadsheetSize)
        {
            return dataItems.Skip(index * maxSpreadsheetSize)
                            .Take(maxSpreadsheetSize)
                            .Where(a => ((object)a) != null)
                            .ToList();
        }

        public Stream EncodeDataToStream(IEnumerable<T> data)
        {
            var stream = new MemoryStream();
            var encodedWorkbook = EncodeDataToWorkbook(data);
            encodedWorkbook.Write(stream);
            return stream;
        }

        public IEnumerable<T> DecodeFileSteam(Stream stream)
        {
            Workbook = new HSSFWorkbook(stream);
            return ConvertToDtoCollection();
        }

        public List<T> DecodeFileToDtoCollection(string filePath)
        {
            List<T> result;

            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None, 4096, true))
            {
                result = DecodeFileSteam(file).ToList();
            }

            return result;
        }

        private IEnumerable<T> ConvertToDtoCollection()
        {
            var result = new List<T>();

            for (var index = 0; index < Workbook.NumberOfSheets; index++)
            {
                var sheet = Workbook.GetSheetAt(index);

                for (var i = 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);

                    if (row == null || !HasCellValues(row)) continue;

                    var rowDto = ConvertRowToDto(row);

                    result.Add(rowDto);
                }
            }
            return result;
        }

        private bool HasCellValues(IRow row)
        {
            return _mapper.Columns.Select(column =>
                row.GetCell(column.Index.GetValueOrDefault()))
                   .Any(cell => cell != null && !string.IsNullOrWhiteSpace(cell.ToString()) && !cell.CellType.Equals(CellType.BLANK));
        }

        private T ConvertRowToDto(IRow row)
        {
            var result = new T();

            var properties = Properties;

            properties.ForEach(property => SetPropertyValue(row, property, result));

            return result;
        }

        private List<PropertyInfo> _properties;

        private List<PropertyInfo> Properties
        {
            get
            {
                return _properties ?? (_properties = typeof(T).GetProperties().ToList());
            }
        }

        private void SetPropertyValue(IRow row, PropertyInfo property, T dto)
        {
            var column = _mapper.Columns.First(c => c.PropertyName == property.Name);

            var cell = row.GetCell(column.Index.GetValueOrDefault());

            var columnType = column.Type;

            if (column.IsNullable)
            {
                if (cell == null || cell.CellType == CellType.BLANK)
                {
                    property.SetValue(dto, null);
                    return;
                }

                columnType = column.UnderLyingType;
            }
            else
            {
                if (cell == null || cell.CellType == CellType.BLANK)
                {
                    throw new Exception(string.Format("A coluna {0} da linha {1} não pode ser nula.", column.PropertyName, (row.RowNum + 1).ToString()));
                }
            }

            if (columnType.IsArray || columnType.GetGenericArguments().Any())
            {
                SetArrayPropertyValue(property, dto, columnType, cell);

                return;
            }

            SetNonArrayPropertyValue(property, dto, columnType, cell);
        }

        private static void SetNonArrayPropertyValue(PropertyInfo property, T dto, Type columnType, ICell cell)
        {
            switch (columnType.Name)
            {
                case "Double":
                case "Single":
                case "Decimal":
                case "Int32":
                    var cellValue = TryGetNumericCellValue(cell);
                    var numericCellValue = Convert.ChangeType(cellValue, columnType);
                    property.SetValue(dto, numericCellValue);
                    break;
                case "DateTime":
                    property.SetValue(dto, cell.DateCellValue);
                    break;
                case "Boolean":
                    SetBooleanPropertyValue(property, dto, cell);
                    break;
                default:
                    property.SetValue(dto, (cell == null ? string.Empty : cell.ToString()));
                    break;
            }
        }

        private static Double TryGetNumericCellValue(ICell cell)
        {
            double numericCellValue;

            try
            {
                numericCellValue = cell.NumericCellValue;
            }
            catch (Exception)
            {
                Double.TryParse(cell.StringCellValue, out numericCellValue);
                if (numericCellValue.ToString() != cell.StringCellValue)
                    throw new Exception(string.Format("Não foi possível transformar em numérico o valor de uma célula do tipo texto. Verifique se o valor da linha {0} coluna {1} está de acordo com o tipo da célula. Valor da celula: {2}", (cell.RowIndex + 1).ToString(), (cell.ColumnIndex + 1).ToString(), cell.StringCellValue));
            }

            return numericCellValue;
        }

        private static void SetArrayPropertyValue(PropertyInfo property, T dto, Type columnType, ICell cell)
        {
            var innerType = columnType.GetElementType() ?? columnType.GetGenericArguments()[0];

            if (cell == null || string.IsNullOrWhiteSpace(cell.ToString()) || cell.CellType == CellType.BLANK)
            {
                property.SetValue(dto, null);
                return;
            }

            var splittedValues = cell.ToString().Split(';')
                .DefaultIfEmpty()
                .Where(value => value != null && !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .ToList();

            switch (innerType.Name)
            {
                case "Double":
                    var doubles = splittedValues.Select(double.Parse);
                    property.SetValue(dto, doubles.ToArray());
                    break;
                case "Single":
                    var floats = splittedValues.Select(float.Parse);
                    property.SetValue(dto, floats.ToArray());
                    break;
                case "Decimal":
                    var decimals = splittedValues.Select(decimal.Parse);
                    property.SetValue(dto, decimals.ToArray());
                    break;
                case "Int32":
                    var ints = splittedValues.Select(int.Parse);
                    property.SetValue(dto, ints.ToArray());
                    break;
                default:
                    property.SetValue(dto, splittedValues.ToArray());
                    break;
            }
        }

        private static void SetBooleanPropertyValue(PropertyInfo property, T dto, ICell cell)
        {
            bool booleanValue;

            if (Boolean.TryParse(cell.ToString(), out booleanValue))
                property.SetValue(dto, booleanValue);
            else
                property.SetValue(dto, null);
        }

        private static HSSFWorkbook InitializeWorkbook()
        {
            var workbook = new HSSFWorkbook();
            //create a entry of DocumentSummaryInformation
            var dsi = PropertySetFactory.CreateDocumentSummaryInformation();

            workbook.DocumentSummaryInformation = dsi;

            //create a entry of SummaryInformation
            var si = PropertySetFactory.CreateSummaryInformation();

            si.Subject = "Default Subject";

            workbook.SummaryInformation = si;

            return workbook;
        }

        public void WriteToFile(string filePath, HSSFWorkbook hssfworkbook)
        {
            using (var sourceStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, true))
            {
                hssfworkbook.Write(sourceStream);
            }
        }

        private void CreateAndPopulateHeader(ISheet sheet)
        {
            var header = sheet.CreateRow(0);
            _mapper.Columns.ForEach(column =>
                                    header.CreateCell(column.Index.GetValueOrDefault(), CellType.STRING)
                                          .SetCellValue(column.HeaderText));
        }

        private void SetColumnDefaults(ISheet sheet)
        {
            var creationHelper = Workbook.GetCreationHelper();

            _mapper.Columns.ForEach(c =>
            {
                var cellStyle = Workbook.CreateCellStyle();
                cellStyle.IsLocked = c.IsReadOnly;
                cellStyle.DataFormat = creationHelper.CreateDataFormat().GetFormat("@");
                sheet.SetDefaultColumnStyle(c.Index.GetValueOrDefault(), cellStyle);
                cellStyle.DataFormat = creationHelper.CreateDataFormat().GetFormat("@");

                var cellType = c.IsNullable ? c.UnderLyingType : c.Type;

                switch (cellType.Name.ToUpper())
                {
                    case "DATETIME":
                        {
                            cellStyle.DataFormat = creationHelper.CreateDataFormat().GetFormat("dd/mm/yyyy");
                            sheet.SetDefaultColumnStyle(c.Index.GetValueOrDefault(), cellStyle);
                        }
                        break;
                    case "STRING":
                        {
                            cellStyle.DataFormat = creationHelper.CreateDataFormat().GetFormat("@");
                            sheet.SetDefaultColumnStyle(c.Index.GetValueOrDefault(), cellStyle);
                        }
                        break;
                    case "DOUBLE":
                    case "SINGLE":
                    case "DECIMAL":
                        {
                            cellStyle.DataFormat = creationHelper.CreateDataFormat().GetFormat("0.000000");
                            sheet.SetDefaultColumnStyle(c.Index.GetValueOrDefault(), cellStyle);
                        }
                        break;
                    case "INT32":
                        {
                            cellStyle.DataFormat = creationHelper.CreateDataFormat().GetFormat("0");
                            sheet.SetDefaultColumnStyle(c.Index.GetValueOrDefault(), cellStyle);
                        }
                        break;
                }

            });
        }

        private void CreateAndPopulateRows(ISheet sheet, IEnumerable<T> dataCollection)
        {
            var rowNum = 1;

            var cachedProperties =
                typeof(T).GetProperties()
                    .Select(x => new { key = x.Name, value = x })
                    .ToDictionary(x => x.key, x => x.value);

            foreach (var dataItem in dataCollection)
            {
                var row = sheet.CreateRow(rowNum);

                _mapper.Columns.ForEach(column =>
                {
                    var cellValue = cachedProperties[column.PropertyName].GetValue(dataItem);

                    var cell = column.CellType == CellType.Unknown
                        ? row.CreateCell(column.Index.GetValueOrDefault())
                        : row.CreateCell(column.Index.GetValueOrDefault(),
                                         column.CellType.GetValueOrDefault(CellType.Unknown));

                    SetCellValue(column, cell, cellValue);
                });

                rowNum++;
            }
        }


        private static void SetCellValue(Column column, ICell cell, object cellValue)
        {
            var columnType = column.Type;

            if (column.IsNullable)
            {
                if (cellValue == null)
                {
                    cell.SetCellValue(string.Empty);
                    cell.SetCellType(CellType.BLANK);
                    return;
                }

                columnType = column.UnderLyingType;
            }

            if (columnType.Name != "String" && cellValue is IEnumerable)
            {
                var values = (cellValue as IEnumerable).Cast<object>();

                cellValue = string.Join(column.ListSeparator, values.Select(value => value.ToString()));
            }

            if (column.CustomEncoder != null)
            {
                cellValue = column.CustomEncoder(cellValue);
            }

            switch (columnType.Name)
            {
                case "DateTime":
                    cell.SetCellValue((DateTime)cellValue);
                    break;
                case "Single":
                case "Decimal":
                case "Double":
                    cell.SetCellValue(cellValue.ToString() == "0" ? 0.0 : Convert.ToDouble(cellValue));
                    cell.SetCellType(CellType.NUMERIC);
                    break;
                case "Int32":
                    cell.SetCellValue((int)cellValue);
                    cell.SetCellType(CellType.NUMERIC);
                    break;
                default:
                    cell.SetCellValue(cellValue == null ? string.Empty : cellValue.ToString());
                    break;
            }

        }

    }

    public static class DataHandlerFactory
    {
        public static DataHandler<T> NewFromColumnMapper<T>(IColumnMapper<T> mapper) where T : new()
        {
            return new DataHandler<T>(mapper);
        }
    }
}
