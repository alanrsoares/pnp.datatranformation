using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NUnit.Framework;
using Newtonsoft.Json;
using Vtex.Practices.DataTransformation.Exceptions;
using Vtex.Practices.DataTransformation.Extensions;
using Vtex.Practices.DataTransformation.ServiceModel;

namespace Vtex.Practices.DataTransformation.Tests
{
    [TestFixture]
    public class Tests
    {
        [Test]
        public void TestAutoMapper()
        {
            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            var properties = mapper.Properties;

            Assert.AreEqual(properties.Count(), mapper.Columns.Count);
        }

        [Test]
        public void TestManualMapper()
        {
            var columnMapper = ColumnMapper<DummyDto>.Factory
                                .CreateNew(true)
                                .Map("Name", "NewName1", false, CellType.STRING, ToUpperCase)
                                .Map("Name", ToUpperCase);

            var properties = typeof(DummyDto).GetProperties();

            Assert.AreEqual(properties.Count(), columnMapper.Columns.Count);
        }

        [Test]
        public void ShouldCreateXlsFileFromDtoCollection_AutoMap()
        {
            var columnMapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            var handler = columnMapper.DataHandler;

            var generatedData = GenerateData();

            var workBook = handler.EncodeDataToWorkbook(generatedData);

            Assert.DoesNotThrow(() => handler.WriteToFile(@"C:\Temp\TestAutomap.xls", workBook));
        }

        [Test]
        public void ShouldCreateXlsFileFromDtoCollection_ManualMap()
        {
            var columnMapper = ColumnMapper<DummyDto>.Factory
                                .CreateNew(true)
                                .Map("Name", "NewName2", false, CellType.STRING, ToUpperCase);

            var handler = columnMapper.DataHandler;
            var generatedData = GenerateData();
            var workBook = handler.EncodeDataToWorkbook(generatedData);

            Assert.DoesNotThrow(() => handler.WriteToFile(@"C:\Temp\TestManualmap.xls", workBook));
        }

        [Test]
        public void ShouldImportFileTest()
        {
            const string filePath = @"C:\Temp\end2endtest.xls";

            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            var handler = mapper.DataHandler;

            var result = handler.DecodeFileToDtoCollection(filePath);

            Assert.IsNotNull(result);
        }

        [Test]
        public void ShouldImportCustomFileTest()
        {
            const string filePath = @"C:\Temp\planilha_productEspec.xls";

            var mapper = ColumnMapper<CustomComplexDto>.Factory
                    .CreateNew(true)
                    .Map("ProductId", "IdProduto (não alterável)")
                    .Map("ProductName", "NomeProduto (não alterável)")
                    .Map("FieldId", "IdCampo (não alterável)")
                    .Map("FieldName", "NomeCampo (não alterável)")
                    .Map("FieldTypeName", "NomeTipoCampo (não alterável)")
                    .Map("FieldValueId", "IdCampoValor (não alterável)")
                    .Map("FieldValueName", "NomeCampoValor (não alterável)")
                    .Map("ProductFieldValueId", "CodigoEspecificaCao (não alterável)")
                    .Map("ProductFieldValueText", "ValorEspecificaCao")
                    .Map("ProductRefId", "CodigoReferencia (não alterável)");

            var handler = mapper.DataHandler;


            IEnumerable<CustomComplexDto> result;

            using (var sourceStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None, 4096, true))
            {
                result = handler.DecodeFileSteam(sourceStream);
            }

            Assert.IsNotNull(handler.Workbook);
            Assert.IsNotNull(result);
        }

        [Test]
        public void ShouldPerformEndToEndTest()
        {
            Console.WriteLine("\nStarting EndToEndTest at {0}", DateTime.Now);

            const string filePath = @"C:\Temp\end2endtest.xls";

            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew().AutoMapColumns();

            var handler = mapper.DataHandler;

            const int dataLength = 1000;

            Console.WriteLine("Generating {0} sample DummyDtos", dataLength);

            var timer = new Stopwatch();

            timer.Start();

            var expectedData = GenerateData(dataLength);

            Console.WriteLine("Generated {0} sample DummyDtos in {1}m {2}s {3}ms",
                dataLength, timer.Elapsed.Minutes, timer.Elapsed.Seconds, timer.Elapsed.Milliseconds);

            Console.WriteLine("Starting data enconding at {0}", DateTime.Now);

            timer.Restart();

            var workBook = handler.EncodeDataToWorkbook(expectedData);

            Console.WriteLine("Encoded {0} sample DummyDtos in {1}m {2}s {3}ms",
                dataLength, timer.Elapsed.Minutes, timer.Elapsed.Seconds, timer.Elapsed.Milliseconds);

            Console.WriteLine("Writing econded workbook to path: {0}", filePath);

            handler.WriteToFile(filePath, workBook);

            Console.WriteLine("Starting data decoding at {0}", DateTime.Now);

            timer.Restart();

            var result = handler.DecodeFileToDtoCollection(filePath);

            Console.WriteLine("Decoded {0} sample DummyDtos in {1}m {2}s {3}ms",
                dataLength, timer.Elapsed.Minutes, timer.Elapsed.Seconds, timer.Elapsed.Milliseconds);

            var serializedResult = JsonConvert.SerializeObject(result);

            var serializedData = JsonConvert.SerializeObject(expectedData);

            Assert.AreEqual(serializedResult, serializedData);
        }

        [Test]
        public void InvalidPropertyMappingTest()
        {
            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            Assert.Throws<InvalidPropertyException>(() => mapper.Map("InvalidPropertyName"));
        }

        [Test]
        public void InvalidColumnMappingTest()
        {
            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            Assert.Throws<InvalidPropertyException>(() => mapper.Map("InvalidPropertyName", ToUpperCase));
        }

        [Test]
        public void ColumnMappingWithColumnObject()
        {
            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew();
            var column = new Column
            {
                PropertyName = "IenumerableOfInts",
                HeaderText = "Header Name",
                ListSeparator = "-"
            };

            mapper.Map(column);

            Assert.AreEqual(1, mapper.Columns.Count);
        }

        [Test]
        public void ColumnMappingWithColumnObjectWithoutHeaderText()
        {
            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew();
            var column = new Column
            {
                PropertyName = "Name"
            };

            mapper.Map(column);

            Assert.AreEqual(column.PropertyName, mapper.Columns[0].HeaderText);
        }

        [Test]
        public void ShouldEncodeXlsStreamTest()
        {
            var data = GenerateData();

            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            var handler = mapper.DataHandler;

            var stream = handler.EncodeToXlsStream(data);

            Assert.IsNotNull(stream);
        }

        [Test]
        public void ShouldEncodeCsvStreamTest()
        {
            var data = GenerateData();

            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            var handler = mapper.DataHandler;

            var stream = handler.EncodeToCsvStream(data);

            Assert.IsNotNull(stream);
        }

        [Test]
        public void ExportHelperXlsTest()
        {
            var data = GenerateData();
            var xlsData = ExportHelper.Export(data).AsXls();
            Assert.IsNotNull(xlsData);
        }

        [Test]
        public void ExportHelperCsvTest()
        {
            var data = GenerateData();
            var csvData = ExportHelper.Export(data, ColumnMapper<DummyDto>.Factory.CreateNew()).AsCsv();
            Assert.IsNull(csvData);
        }

        [Test]
        public void ParameterTypeInferenceAndCustomTransformTest()
        {
            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew();
            mapper.Map("Name", ToUpperCase);

            var mappedColumn = mapper.Columns.FirstOrDefault();

            Assert.IsNotNull(mappedColumn);
            Assert.AreEqual(mappedColumn.PropertyName, mappedColumn.HeaderText);
            Assert.AreEqual(CellType.STRING, mappedColumn.CellType);
        }

        [Test]
        public void ParameterTypeInferenceTest()
        {
            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew();
            mapper.Map("Name");

            var mappedColumn = mapper.Columns.FirstOrDefault();

            Assert.IsNotNull(mappedColumn);
            Assert.AreEqual("Name", mappedColumn.PropertyName);
            Assert.AreEqual(mappedColumn.PropertyName, mappedColumn.HeaderText);
            Assert.AreEqual(CellType.STRING, mappedColumn.CellType);
        }

        [Test]
        public void InvalidColumnMappingWithComplexParameters()
        {
            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew();

            Assert.Throws<InvalidPropertyException>(() => mapper.Map("InvalidPropertyName"));
        }

        [Test]
        public void ShouldPerformSelfTest()
        {
            const string input = "Hello";
            var output = ToUpperCase(input);

            const string input2 = "one,two";
            var output2 = CommaToDashReplacer(input2);

            Assert.AreEqual(input.ToUpper(), output);
            Assert.AreEqual(input2.Replace(",", "-"), output2);
        }

        [Test]
        public void ShouldMapTwoColumnAndUnmapTheFirstByPropertyName()
        {
            var propertyNames = new[] { "Salary", "Id" };

            var mapper = ColumnMapper<DummyDto>.Factory
                .CreateNew()
                .Map(propertyNames)
                .Unmap(propertyNames[0]);

            Assert.AreEqual(mapper.Columns.Count, 1);
            Assert.AreEqual(mapper.Columns[0].PropertyName, propertyNames[1]);
        }

        [Test]
        public void ShouldUnmapEmptyNameWithErrors()
        {
            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            Assert.Throws<InvalidPropertyException>(() => mapper.Unmap(string.Empty));
        }

        [Test]
        public void ShouldUnmapMultipleProperties()
        {
            var propertyNames = new[] { "Salary", "Id" };

            var mapper = ColumnMapper<DummyDto>.Factory
                .CreateNew()
                .Map(propertyNames)
                .Unmap(propertyNames);

            Assert.AreEqual(mapper.Columns.Count, 0);
        }

        [Test]
        public void ShouldMapColumnsUsingLambdaExpression()
        {
            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew()
                            .Map(x => x.Name)
                            .Map(x => x.Id)
                            .Map(x => x.Salary);

            Assert.AreEqual(mapper.Columns.Count, 3);
        }

        [Test]
        public void ShouldMapColumnsWithArbitraryCellTypeUsingLambdaExpression()
        {
            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew()
                            .Map(x => x.Name, CellType.STRING)
                            .Map(x => x.Id, CellType.STRING)
                            .Map(x => x.Salary, CellType.STRING);

            Assert.AreEqual(mapper.Columns.Count, 3);
            Assert.AreEqual(mapper.Columns[0].CellType, CellType.STRING);
            Assert.AreEqual(mapper.Columns[1].CellType, CellType.STRING);
            Assert.AreEqual(mapper.Columns[2].CellType, CellType.STRING);
        }

        [Test]
        public void ShouldExportWithReadOnlyColumn()
        {
            const string filePath = @"C:\Temp\readonly.xls";

            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew()
                            .Map(x => x.Name, "Name (ReadOnly)", true)
                            .Map(x => x.Id, "Id")
                            .Map(x => x.Salary, "Salary");
            var handler = mapper.DataHandler;

            var expected = GenerateData(30);

            var wb = handler.EncodeDataToWorkbook(expected);

            handler.WriteToFile(filePath, wb);

            Assert.IsNotNull(wb);
        }


        private static IEnumerable<DummyDto> GenerateData(int length = 500)
        {
            return DummyDtoFactory.GenerateData(length).ToList();
        }

        private static object ToUpperCase(object value)
        {
            return value.ToString().ToUpper();
        }

        private static object CommaToDashReplacer(object value)
        {
            return value.ToString().Replace(',', '-');
        }
    }

    public class CustomComplexDto
    {
        public int? ProductId { get; set; }
        public string ProductName { get; set; }
        public int? FieldId { get; set; }
        public string FieldName { get; set; }
        public string FieldTypeName { get; set; }
        public string FieldValueId { get; set; }
        public string FieldValueName { get; set; }
        public string ProductFieldValueId { get; set; }
        public string ProductFieldValueText { get; set; }
        public string ProductRefId { get; set; }
    }
}