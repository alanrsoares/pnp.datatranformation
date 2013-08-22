using System;

namespace Vtex.Practices.DataTransformation.Tests
{
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