using System;
using System.Collections.Generic;
using NPOI.SS.Formula.Functions;

namespace Vtex.Practices.DataTransformation.Tests
{
    public class DummyDto
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime BirthDate { get; set; }
        public bool IsActive { get; set; }
        public decimal? Salary { get; set; }
        public int? NullableIntegerValue { get; set; }
        public DateTime? NullableDateTimeValue { get; set; }
        public Single SingleValue { get; set; }
        public Double DoubleValue { get; set; }
        public IEnumerable<Int32> IenumerableOfInts { get; set; }
        public IEnumerable<Double> IenumerableOfDoubles { get; set; }
        public IEnumerable<Single> IenumerableOfFloats { get; set; }
        public IEnumerable<Decimal> IenumerableOfDecimals { get; set; }
        public int[] ArrayOfInts { get; set; }
        public IEnumerable<string> IenumerableOfStrings { get; set; }
        public string[] ArrayOfStrings { get; set; }
        public IEnumerable<int> NullEnumerableOfInts { get; set; }
    }
}