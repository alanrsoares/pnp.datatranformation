using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Vtex.Practices.DataTransformation.xls;

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

            CreateAndPopulateRows(sheet, data);

            _mapper.Columns.ForEach(c => sheet.AutoSizeColumn(c.Index));

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

            var cell = row.GetCell(column.Index);

            var underlyingType = Nullable.GetUnderlyingType(column.Type);

            var columnType = column.Type;

            if (underlyingType != null)
            {
                if (cell == null || cell.CellType == CellType.BLANK)
                {
                    property.SetValue(dto, null);
                    return;
                }

                columnType = underlyingType;
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

            ////create a entry of DocumentSummaryInformation
            var dsi = PropertySetFactory.CreateDocumentSummaryInformation();

            workbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
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
                                    header.CreateCell(column.Index)
                                          .SetCellValue(column.HeaderText));
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

                            var cell = column.CellType == CellType.Unknown ? row.CreateCell(column.Index) : row.CreateCell(column.Index, column.CellType.GetValueOrDefault(CellType.Unknown));

                            SetCellValue(column, cell, cellValue);
                        });

                    rowNum++;
                });
        }

        private void SetCellValue(Column column, ICell cell, object cellValue)
        {
            var dateCellStyle = Workbook.CreateCellStyle();
            var creationHelper = Workbook.GetCreationHelper();

            dateCellStyle.DataFormat = creationHelper.CreateDataFormat().GetFormat("dd/mm/yyyy");

            var underlyingType = Nullable.GetUnderlyingType(column.Type);

            var columnType = column.Type;

            if (underlyingType != null)
            {
                if (cellValue == null)
                {
                    cell.SetCellValue(string.Empty);
                    cell.SetCellType(CellType.BLANK);
                    return;
                }

                columnType = underlyingType;
            }

            if (column.CustomTransformAction != null)
            {
                cellValue = column.CustomTransformAction(cellValue);
            }

            switch (columnType.Name)
            {
                case "DateTime":
                    cell.SetCellValue((DateTime)cellValue);
                    cell.CellStyle = dateCellStyle;
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
                    cell.SetCellValue(cellValue.ToString());
                    break;
            }

        }

    }
}