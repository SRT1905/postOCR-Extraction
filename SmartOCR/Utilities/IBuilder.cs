namespace SmartOCR
{
    /// <summary>
    /// Defines a generic interface for any class implementing Builder pattern.
    /// </summary>
    /// <typeparam name="T">Type of instance being built.</typeparam>
    public interface IBuilder<out T>
        where T : class
    {
        /// <summary>
        /// Builds an instance of <typeparamref name="T"/> class.
        /// </summary>
        /// <returns>An instance of <typeparamref name="T"/> class.</returns>
        T Build();

        /// <summary>
        /// Sets a new instance of <typeparamref name="T"/> class to build.
        /// </summary>
        void Reset();
    }
}
