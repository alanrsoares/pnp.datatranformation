namespace Vtex.Practices.DataTransformation
{
    public class ColumnMapperFactory<T> where T : new()
    {
        public ColumnMapper<T> CreateNew()
        {
            return new ColumnMapper<T>();
        }

        public ColumnMapper<T> CreateNew(bool autoMap)
        {
            var mapper = CreateNew();

            if (autoMap) mapper.AutoMapColumns();

            return mapper;
        }
    }
}