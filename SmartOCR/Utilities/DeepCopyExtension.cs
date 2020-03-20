namespace SmartOCR.Utilities
{
    using System.IO;
    using System.Runtime.Serialization.Formatters.Binary;
    using System.Xml.Serialization;

    /// <summary>
    /// Contains generic extension methods.
    /// </summary>
    public static class DeepCopyExtension
    {
        /// <summary>
        /// Performs deep cloning of provided generic object.
        /// An object must have a parameterless constructor.
        /// </summary>
        /// <typeparam name="T">Type of cloned object.</typeparam>
        /// <param name="source">An object to be cloned.</param>
        /// <returns>A deep copy of current object.</returns>
        public static T DeepCopyXml<T>(this T source)
        {
            var serializer = new XmlSerializer(typeof(T));
            using (var memoryStream = new MemoryStream())
            {
                serializer.Serialize(memoryStream, source);
                memoryStream.Position = 0;
                return (T)serializer.Deserialize(memoryStream);
            }
        }

        /// <summary>
        /// Performs deep cloning of provided generic object.
        /// An object and its internal property hierarchy must have a <see cref="System.SerializableAttribute"/> attribute.
        /// </summary>
        /// <typeparam name="T">Type of cloned object.</typeparam>
        /// <param name="source">An object to be cloned.</param>
        /// <returns>A deep copy of current object.</returns>
        public static T DeepCopy<T>(this T source)
        {
            using (var memoryStream = new MemoryStream())
            {
                var binaryFormatter = new BinaryFormatter();
                binaryFormatter.Serialize(memoryStream, source);
                memoryStream.Seek(0, SeekOrigin.Begin);
                return (T)binaryFormatter.Deserialize(memoryStream);
            }
        }
    }
}