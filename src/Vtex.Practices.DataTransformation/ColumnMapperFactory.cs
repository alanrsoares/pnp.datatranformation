namespace Vtex.Practices.DataTransformation
{
    public class ColumnMapperFactory<T> where T : new()
    {
        public IColumnMapper<T> CreateNew()
        {
            return new ColumnMapper<T>();
        }

        public IColumnMapper<T> CreateNew(bool autoMap)
        {
            var mapper = CreateNew();

            if (autoMap) mapper.AutoMapColumns();

            return mapper;
        }
    }
}