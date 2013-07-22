using System;
using System.Collections.Generic;
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
            var columnMapper = new ColumnMapper<DummyDto>();

            columnMapper.AutoMapColumns();

            Assert.AreEqual(6, columnMapper.Columns.Count);
        }

        [Test]
        public void TestManualMapper()
        {
            var columnMapper = ManualMapper();

            Assert.AreEqual(6, columnMapper.Columns.Count);
        }

        [Test]
        public void TestXlsCreationFromDtoCollection_AutoMap()
        {
            var columnMapper = AutoMapper();

            var handler = new DataHandler<DummyDto>(columnMapper);
            var workBook = handler.EncodeDataToWorkbook(Data);

            handler.WriteToFile(@"C:\Temp\TestAutomap.xls", workBook);
        }

        [Test]
        public void TestXlsCreationFromDtoCollection_ManualMap()
        {
            var columnMapper = ManualMapper();

            var handler = new DataHandler<DummyDto>(columnMapper);
            var workBook = handler.EncodeDataToWorkbook(Data);

            handler.WriteToFile(@"C:\Temp\TestManualmap.xls", workBook);
        }

        [Test]
        public void ImportFile()
        {
            const string filePath = @"C:\Temp\TestAutomap.xls";

            var mapper = AutoMapper();
            var handler = new DataHandler<DummyDto>(mapper);

            var result = handler.DecodeFileToDtoCollection(filePath);

            Assert.IsNotNull(result);
        }

        [Test]
        public void EndToEndTest()
        {
            const string filePath = @"C:\Temp\end2endtest.xls";

            var mapper = AutoMapper();

            var handler = new DataHandler<DummyDto>(mapper);

            var workBook = handler.EncodeDataToWorkbook(Data);

            handler.WriteToFile(filePath, workBook);

            var result = handler.DecodeFileToDtoCollection(filePath);

            var serializedResult = JsonConvert.SerializeObject(result);

            var serializedData = JsonConvert.SerializeObject(Data);

            Assert.AreEqual(serializedResult, serializedData);
        }

        private static ColumnMapper<DummyDto> AutoMapper()
        {
            var columnMapper = new ColumnMapper<DummyDto>();

            columnMapper.AutoMapColumns();

            return columnMapper;
        }

        private static ColumnMapper<DummyDto> ManualMapper()
        {
            var columnMapper = new ColumnMapper<DummyDto>();

            columnMapper
                .MapColumn("Id", "Id", CellType.NUMERIC)
                .MapColumn("Name", "NewName", CellType.STRING, value => value.ToString().ToUpper())
                .MapColumn("BirthDate")
                .MapColumn("NullableIntegerValue", CellType.NUMERIC)
                .MapColumn("NullableDateTimeValue")
                .MapColumn(1, "Name", "Overloaded Column Name", CellType.STRING, value => value.ToString().Replace(',', '-'))
                .MapColumn("Salary");

            return columnMapper;
        }

        public List<DummyDto> Data
        {
            get
            {
                return new List<DummyDto>
                    {
                        new DummyDto{Id = 1,Name = "Name 1", BirthDate = new DateTime(1987,7,7)},
                        new DummyDto{Id = 2,Name = "2", BirthDate = new DateTime(1987,7,7),NullableIntegerValue = 135},
                        new DummyDto{Id = 3,Name = "Name 3", BirthDate = new DateTime(1987,7,7),NullableDateTimeValue = new DateTime(2013,9,18)},
                        new DummyDto{Id = 1,Name = "Name 1", BirthDate = new DateTime(1987,7,7)},
                        new DummyDto{Id = 2,Name = "2,1,3,5,7,8", BirthDate = new DateTime(1987,7,7),NullableIntegerValue = 135},
                        new DummyDto{Id = 3,Name = "Name 3", BirthDate = new DateTime(1987,7,7),NullableDateTimeValue = new DateTime(2013,9,18)},
                        new DummyDto{Id = 1,Name = "Name 1", BirthDate = new DateTime(1987,7,7)},
                        new DummyDto{Id = 2,Name = "2", BirthDate = new DateTime(1987,7,7),NullableIntegerValue = 135},
                        new DummyDto{Id = 3,Name = "Name 3", BirthDate = new DateTime(1987,7,7),NullableDateTimeValue = new DateTime(2013,9,18)},
                        new DummyDto{Id = 1,Name = "Name 1", BirthDate = new DateTime(1987,7,7)},
                        new DummyDto{Id = 2,Name = "2", BirthDate = new DateTime(1987,7,7),NullableIntegerValue = 135},
                        new DummyDto{Id = 3,Name = "Name 3", BirthDate = new DateTime(1987,7,7),NullableDateTimeValue = new DateTime(2013,9,18)},
                        new DummyDto{Id = 1,Name = "Name 1", BirthDate = new DateTime(1987,7,7)},
                        new DummyDto{Id = 2,Name = "2", BirthDate = new DateTime(1987,7,7),NullableIntegerValue = 135},
                        new DummyDto{Id = 3,Name = "Name 3", BirthDate = new DateTime(1987,7,7),NullableDateTimeValue = new DateTime(2013,9,18)},
                        new DummyDto{Id = 1,Name = "Name 1", BirthDate = new DateTime(1987,7,7)},
                        new DummyDto{Id = 2,Name = "2", BirthDate = new DateTime(1987,7,7),NullableIntegerValue = 135},
                        new DummyDto{Id = 3,Name = "Name 3", BirthDate = new DateTime(1987,7,7),NullableDateTimeValue = new DateTime(2013,9,18)},
                        new DummyDto{Id = 1,Name = "Name 1", BirthDate = new DateTime(1987,7,7), Salary = new decimal(3500.76)},
                        new DummyDto{Id = 2,Name = "2", BirthDate = new DateTime(1987,7,7),NullableIntegerValue = 135},
                        new DummyDto{Id = 3,Name = "Name 3", BirthDate = new DateTime(1987,7,7),NullableDateTimeValue = new DateTime(2013,9,18)},
                        new DummyDto{Id = 1,Name = "Name 1", BirthDate = new DateTime(1987,7,7)},
                        new DummyDto{Id = 2,Name = "2", BirthDate = new DateTime(1987,7,7),NullableIntegerValue = 135},
                        new DummyDto{Id = 3,Name = "Name 3", BirthDate = new DateTime(1987,7,7),NullableDateTimeValue = new DateTime(2013,9,18)},
                        new DummyDto{Id = 1,Name = "Name 1", BirthDate = new DateTime(1987,7,7)},
                        new DummyDto{Id = 2,Name = "2", BirthDate = new DateTime(1987,7,7),NullableIntegerValue = 135},
                        new DummyDto{Id = 3,Name = "Name 3", BirthDate = new DateTime(1987,7,7),NullableDateTimeValue = new DateTime(2013,9,18)},
                    };
            }
        }

    }

    public class DummyDto
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime BirthDate { get; set; }
        public decimal? Salary { get; set; }
        public int? NullableIntegerValue { get; set; }
        public DateTime? NullableDateTimeValue { get; set; }
    }
}