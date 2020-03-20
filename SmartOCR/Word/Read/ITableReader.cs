namespace SmartOCR.Word.Read
{
    using System.Collections.Generic;

    /// <summary>
    /// Defines a generalized interface for any class performing collection and transformation of Word tables.
    /// </summary>
    public interface ITableReader
    {
        /// <summary>
        /// Returns a collection of transformed Word tables within page provided in class constructor.
        /// </summary>
        /// <returns>An instance of <see cref="List{WordTable}"/>.</returns>
        List<WordTable> GetWordTables();
    }
}