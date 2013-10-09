namespace Vtex.Practices.DataTransformation
{
    public class ColumnMapperFactory<T> where T : new()
    {
        public IColumnMapper<T> CreateNew(bool autoMap = false)
        {
            var mapper = new ColumnMapper<T>();

            if (autoMap) mapper.AutoMapColumns();

            return mapper;
        }
    }
}