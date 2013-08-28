using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using NUnit.Framework;
using Newtonsoft.Json;

namespace Vtex.Practices.DataTransformation.Tests
{
    [TestFixture]
    public class Tests
    {
        [Test]
        public void TestAutoMapper()
        {
            var columnMapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            var properties = typeof(DummyDto).GetProperties();

            Assert.AreEqual(properties.Count(), columnMapper.Columns.Count);
        }

        [Test]
        public void TestManualMapper()
        {
            var columnMapper = ColumnMapper<DummyDto>.Factory
                                .CreateNew(true)
                                .MapColumn("Name", "NewName1", CellType.STRING, value => value.ToString().ToUpper())
                                .MapColumn(1, value => value.ToString().Replace(',', '-'));

            var properties = typeof(DummyDto).GetProperties();

            Assert.AreEqual(properties.Count(), columnMapper.Columns.Count);
        }

        [Test]
        public void TestXlsCreationFromDtoCollection_AutoMap()
        {
            var columnMapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            var handler = columnMapper.DataHandler;

            var generatedData = GenerateData();

            var workBook = handler.EncodeDataToWorkbook(generatedData);

            Assert.DoesNotThrow(() => handler.WriteToFile(@"C:\Temp\TestAutomap.xls", workBook));
        }

        [Test]
        public void TestXlsCreationFromDtoCollection_ManualMap()
        {
            var columnMapper = ColumnMapper<DummyDto>.Factory
                                .CreateNew(true)
                                .MapColumn("Name", "NewName2", CellType.STRING, value => value.ToString().ToUpper())
                                .MapColumn(1, value => value.ToString().Replace(',', '-'))
                                .MapColumn(0, "NewColumnName");

            var handler = columnMapper.DataHandler;
            var generatedData = GenerateData();
            var workBook = handler.EncodeDataToWorkbook(generatedData);

            Assert.DoesNotThrow(() => handler.WriteToFile(@"C:\Temp\TestManualmap.xls", workBook));
        }

        [Test]
        public void ImportFile()
        {
            const string filePath = @"C:\Temp\TestAutomap.xls";

            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            var handler = mapper.DataHandler;

            var result = handler.DecodeFileToDtoCollection(filePath);

            Assert.IsNotNull(result);
        }

        [Test]
        public void ImportCustomFile()
        {
            const string filePath = @"C:\Temp\planilha_productEspec.xls";

            var mapper = ColumnMapper<ProductsSpecificationByCategoryIdDto>.Factory
                    .CreateNew(true)
                    .MapColumn("ProductId", "IdProduto (não alterável)")
                    .MapColumn("ProductName", "NomeProduto (não alterável)")
                    .MapColumn("FieldId", "IdCampo (não alterável)")
                    .MapColumn("FieldName", "NomeCampo (não alterável)")
                    .MapColumn("FieldTypeName", "NomeTipoCampo (não alterável)")
                    .MapColumn("FieldValueId", "IdCampoValor (não alterável)")
                    .MapColumn("FieldValueName", "NomeCampoValor (não alterável)")
                    .MapColumn("ProductFieldValueId", "CodigoEspecificaCao (não alterável)")
                    .MapColumn("ProductFieldValueText", "ValorEspecificaCao")
                    .MapColumn("ProductRefId", "CodigoReferencia (não alterável)");

            var handler = new DataHandler<ProductsSpecificationByCategoryIdDto>(mapper);

            var result = handler.DecodeFileToDtoCollection(filePath);

            Assert.IsNotNull(result);
        }

        [Test]
        public void EndToEndTest()
        {
            const string filePath = @"C:\Temp\end2endtest.xls";

            var mapper = ColumnMapper<DummyDto>.Factory.CreateNew(true);

            var handler = mapper.DataHandler;

            var expectedData = GenerateData();

            var workBook = handler.EncodeDataToWorkbook(expectedData);

            handler.WriteToFile(filePath, workBook);

            var result = handler.DecodeFileToDtoCollection(filePath);

            var serializedResult = JsonConvert.SerializeObject(result);

            var serializedData = JsonConvert.SerializeObject(expectedData);

            Assert.AreEqual(serializedResult, serializedData);
        }

        public List<DummyDto> GenerateData()
        {
            return DummyDtoFactory.NewDummyDtos(100).ToList();
        }

    }

    public class ProductsSpecificationByCategoryIdDto
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