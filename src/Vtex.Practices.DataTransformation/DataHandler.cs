using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Vtex.Practices.DataTransformation.ServiceModel;

namespace Vtex.Practices.DataTransformation
{
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

            var sheet = Workbook.CreateSheet();

            CreateAndPopulateHeader(sheet);

            SetColumnDefaults(sheet);

            CreateAndPopulateRows(sheet, data);

            _mapper.Columns.ForEach(c => sheet.AutoSizeColumn(c.Index.GetValueOrDefault()));

            return Workbook;
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

            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                result = DecodeFileSteam(file).ToList();
            }

            return result;
        }

        private IEnumerable<T> ConvertToDtoCollection()
        {
            var sheet = Workbook.GetSheetAt(0);

            var result = new List<T>();

            for (var i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);

                var rowDto = ConvertRowToDto(row);

                result.Add(rowDto);
            }

            return result;
        }

        private T ConvertRowToDto(IRow row)
        {
            var result = new T();

            var properties = typeof(T).GetProperties().ToList();

            properties.ForEach(property =>
                                SetPropertyValue(row, property, result));

            return result;
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

            if (columnType.IsArray || columnType.GetGenericArguments().Any())
            {
                var innerType = columnType.GetElementType() ?? columnType.GetGenericArguments()[0];

                var splittedValues = cell.ToString().Split(';')
                                                    .Where(value => !string.IsNullOrWhiteSpace(value))
                                                    .Select(value => value.Trim())
                                                    .ToList();

                if (!splittedValues.Any() || cell.CellType == CellType.BLANK)
                {
                    property.SetValue(dto, null);
                    return;
                }

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

                return;
            }

            switch (columnType.Name)
            {
                case "DateTime":
                    property.SetValue(dto, cell.DateCellValue);
                    break;
                case "Double":
                    property.SetValue(dto, cell.NumericCellValue);
                    break;
                case "Single":
                    property.SetValue(dto, (float)cell.NumericCellValue);
                    break;
                case "Decimal":
                    property.SetValue(dto, (decimal)cell.NumericCellValue);
                    break;
                case "Int32":
                    property.SetValue(dto, (Int32)cell.NumericCellValue);
                    break;
                default:
                    property.SetValue(dto, (cell == null ? string.Empty : cell.ToString()));
                    break;
            }
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
            //Write the stream data of workbook to the root directory
            var file = new FileStream(filePath, FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
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
                }

            });
        }

        private void CreateAndPopulateRows(ISheet sheet, IEnumerable<T> dataSet)
        {
            var rowNum = 1;
            dataSet.ToList().ForEach(dto =>
                {
                    var row = sheet.CreateRow(rowNum);

                    _mapper.Columns.ForEach(column =>
                        {
                            var cellValue = dto.GetType().GetProperty(column.PropertyName).GetValue(dto);

                            var cell = column.CellType == CellType.Unknown
                                ? row.CreateCell(column.Index.GetValueOrDefault())
                                : row.CreateCell(column.Index.GetValueOrDefault(), column.CellType.GetValueOrDefault(CellType.Unknown));

                            SetCellValue(column, cell, cellValue);
                        });

                    rowNum++;
                });
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
                    cell.SetCellValue(cellValue.ToString() == "0" ? 0.0 : Convert.ToSingle(cellValue));
                    break;
                case "Decimal":
                case "Double":
                    cell.SetCellValue(cellValue.ToString() == "0" ? 0.0 : Convert.ToDouble(cellValue));
                    break;
                case "Int32":
                    cell.SetCellValue((int)cellValue);
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