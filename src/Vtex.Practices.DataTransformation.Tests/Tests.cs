using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
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

            Assert.DoesNotThrow(() => handler.WriteToFile(@"C:\Temp\TestAutomap.xls", workBook));
        }

        [Test]
        public void TestXlsCreationFromDtoCollection_ManualMap()
        {
            var columnMapper = ManualMapper();

            var handler = new DataHandler<DummyDto>(columnMapper);
            var workBook = handler.EncodeDataToWorkbook(Data);

            Assert.DoesNotThrow(() => handler.WriteToFile(@"C:\Temp\TestManualmap.xls", workBook));
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

            var expectedData = Data;

            var workBook = handler.EncodeDataToWorkbook(expectedData);

            handler.WriteToFile(filePath, workBook);

            var result = handler.DecodeFileToDtoCollection(filePath);

            var serializedResult = JsonConvert.SerializeObject(result);

            var serializedData = JsonConvert.SerializeObject(expectedData);

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
                var dtos = new List<DummyDto>();

                while (dtos.Count < 50)
                {
                    dtos.Add(DummyDtoFactory.NewDummyDto());
                }

                return dtos;
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

    public static class DummyDtoFactory
    {
        public static DummyDto NewDummyDto()
        {
            return new DummyDto
                {
                    Id = GetRandom(0, 100),
                    Name = GetRamdomName(),
                    BirthDate = GetRandomDateTime(),
                    Salary = new decimal(GetRandom(4000, 10000) + GetRandom(100, 1000) * .1),
                    NullableIntegerValue = new int?[] { GetRandom(), null, null }[GetRandom(0, 2)],
                    NullableDateTimeValue = new DateTime?[] { GetRandomDateTime(), null, null }[GetRandom(0, 2)],
                };
        }

        private static string GetRamdomName()
        {
            return string.Format("{0} {1}",
                            new[] { "John", "Jack", "Juca", "Chico", "Zé" }[GetRandom()],
                            new[] { "Doe", "Ellis", "The Ripper", "Roberto", "Johnson" }[GetRandom()]);
        }

        private static Random GetRandomizer()
        {
            var newGuid = Guid.NewGuid();
            var guidNumber = Regex.Replace(newGuid.ToString(), @"[^\d]", "");
            var rnd0 = new Random(Convert.ToInt32(guidNumber.Substring(0, 8)));
            var rnd = new Random(rnd0.Next(5000));
            return rnd;
        }

        private static DateTime GetRandomDateTime()
        {
            var year = GetRandom(1960, 1990);
            var month = GetRandom(1, 12);
            var day = GetRandom(1, month == 2 ? 29 : 30);
            var randomDateTime = new DateTime(year, month, day);

            return randomDateTime;
        }

        private static int GetRandom(int min = 0, int max = 4)
        {
            var rnd = GetRandomizer();
            return rnd.Next(min, max);
        }
    }
}