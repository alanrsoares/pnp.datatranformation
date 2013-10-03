using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Vtex.Practices.DataTransformation.Tests
{
    public static class DummyDtoFactory
    {
        public static DummyDto NewDummyDto()
        {
            return new DummyDto
                {
                    Id = GetRandom(0, 100),
                    Name = GetRamdomName(),
                    BirthDate = GetRandomDateTime(),
                    Salary = new decimal(GetRandom(4000, 10000) + (GetRandom(1, 100) * .01)),
                    NullableIntegerValue = new int?[] { GetRandom(), null, null }[GetRandom(0, 2)],
                    NullableDateTimeValue = new DateTime?[] { GetRandomDateTime(), null, null }[GetRandom(0, 2)],
                    IenumerableOfInts = new[] { 1, 2, 3, 4, 5 },
                    ArrayOfInts = new[] { 1, 2, 3, 4, 5 },
                    IenumerableOfStrings = new[] { "ab", "cd", "ef" },
                    ArrayOfStrings = new[] { "ab", "cd", "ef" }
                };
        }

        private static string GetRamdomName()
        {
            return string.Format("{0} {1}",
                                 new[] { "Marshall", "Jack", "Lucy", "John", "Joe" }[GetRandom()],
                                 new[] { "Doe", "Ellis", "The Ripper", "Erikssen", "Johnson" }[GetRandom()]);
        }

        private static Random GetRandomizer()
        {
            var newGuid = Guid.NewGuid();
            var guidNumber = Regex.Replace(newGuid.ToString(), @"[^\d]", "");
            var rnd0 = new Random(Convert.ToInt32(guidNumber.Substring(0, 9)));
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

        private static int GetRandom(int min = 0, int max = 5)
        {
            var rnd = GetRandomizer();
            return rnd.Next(min, max);
        }

        public static IEnumerable<DummyDto> GenerateData(int amount)
        {
            var dtos = new List<DummyDto>();

            while (dtos.Count < 50)
            {
                dtos.Add(NewDummyDto());
            }

            return dtos;
        }
    }
}