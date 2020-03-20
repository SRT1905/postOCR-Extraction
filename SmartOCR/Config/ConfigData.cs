namespace SmartOCR.Config
{
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Used as a container of config data from Excel workbook.
    /// </summary>
    public class ConfigData
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigData"/> class.
        /// Instance has empty config field collection.
        /// </summary>
        public ConfigData()
        {
            this.Fields = new List<ConfigField>();
        }

        /// <summary>
        /// Gets collection of config fields.
        /// </summary>
        public List<ConfigField> Fields { get; }

        /// <summary>
        /// Gets config field at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of element to get.</param>
        /// <returns>Single config field.</returns>
        public ConfigField this[int index] => this.Fields[index];

        /// <summary>
        /// Gets config field by the specified string index.
        /// </summary>
        /// <param name="name">String representation of element to get.</param>
        /// <returns>Single config field.</returns>
        public ConfigField this[string name]
        {
            get
            {
                return this.Fields.First(singleField => singleField.Name == name);
            }
        }

        /// <summary>
        /// Inserts <see cref="ConfigField"/> instance to field collection.
        /// </summary>
        /// <param name="field"><see cref="ConfigField"/> instance that describes single search field.</param>
        public void AddField(ConfigField field)
        {
            if (field != null)
            {
                this.Fields.Add(field);
            }
        }
    }
}
