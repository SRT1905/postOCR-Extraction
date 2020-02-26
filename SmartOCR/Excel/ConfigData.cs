﻿using System;
using System.Collections.Generic;
using System.Linq;

namespace SmartOCR
{
    /// <summary>
    /// Used as a container of config data from Excel workbook.
    /// </summary>
    public class ConfigData
    {
        #region Properties
        /// <summary>
        /// Collection of config fields.
        /// </summary>
        public List<ConfigField> Fields { get; }
        #endregion

        #region Indexers
        /// <summary>
        /// Gets config field at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of element to get.</param>
        /// <returns>Single config field.</returns>
        public ConfigField this[int index]
        {
            get
            {
                return Fields[index];
            }
        }
        /// <summary>
        /// Gets config field by the specified string index.
        /// </summary>
        /// <param name="name">String representation of element to get.</param>
        /// <returns>Single config field.</returns>
        public ConfigField this[string name]
        {
            get
            {
                return Fields.Where(singleField => singleField.Name == name).First();
            }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new <see cref="ConfigData"/> instance with empty config fields collection.
        /// </summary>
        public ConfigData()
        {
            Fields = new List<ConfigField>();
        }
        #endregion

        #region Public methods
        /// <summary>
        /// Inserts <see cref="ConfigField"/> instance to field collection.
        /// </summary>
        /// <param name="field"><see cref="ConfigField"/> instance that describes single search field.</param>
        /// <exception cref="ArgumentNullException"></exception>
        public void AddField(ConfigField field)
        {
            if (field != null)
            {
                Fields.Add(field);
            }
            //else
            //{
            //    throw new ArgumentNullException(nameof(field), Properties.Resources.nullConfigField);
            //}
        }
        #endregion
    }
}
