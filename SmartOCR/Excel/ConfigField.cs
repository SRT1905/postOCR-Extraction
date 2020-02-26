using System;
using System.Collections.Generic;
using System.Linq;

namespace SmartOCR
{
    /// <summary>
    /// Describes single search field defined in Excel config file.
    /// </summary>
    public class ConfigField
    {
        #region Properties
        /// <summary>
        /// Collection of search expressions that define search process for current field.
        /// </summary>
        public List<ConfigExpression> Expressions { get; } = new List<ConfigExpression>();
        /// <summary>
        /// String value that is used to check similarity between it and values that match <see cref="RegExPattern"/>.
        /// </summary>
        public string ExpectedName { get; set; }
        /// <summary>
        /// Regular expression pattern.
        /// </summary>
        public string RegExPattern { get; set; }
        /// <summary>
        /// Field name.
        /// </summary>
        public string Name { get; }
        /// <summary>
        /// Data type of searched value.
        /// </summary>
        public string ValueType { get; }
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new <see cref="ConfigField"/> instance with name and value type.
        /// </summary>
        /// <param name="name">Field name.</param>
        /// <param name="valueType">Data type of searched value.</param>
        public ConfigField(string name, string valueType)
        {
            Name = name;
            ValueType = valueType;
        }
        #endregion

        #region Public methods
        /// <summary>
        /// Parses Excel cell contents and gets regular expression pattern and expected field name.
        /// </summary>
        /// <param name="input">String representation of Excel cell contents.</param>
        public void ParseFieldExpression(string input)
        {
            if (input == null)
            {
                RegExPattern = Name;
                ExpectedName = Name;
                return;
            }
            var splittedValue = input.Split(';').ToList();

            while (splittedValue.Count != 2)
            {
                splittedValue[0] = $"{splittedValue[0]};{splittedValue[1]}";
                for (int i = 2; i < splittedValue.Count; i++)
                {
                    splittedValue[i - 1] = splittedValue[i];
                }
                splittedValue.RemoveAt(splittedValue.Count - 1);
            }

            RegExPattern = string.IsNullOrEmpty(splittedValue[0])
                ? Name
                : splittedValue[0];
            ExpectedName = string.IsNullOrEmpty(splittedValue[1])
                ? Name
                : splittedValue[1];
        }
        /// <summary>
        /// Inserts <see cref="ConfigExpression"/> instance to expression collection.
        /// </summary>
        /// <param name="expression"><see cref="ConfigExpression"/> instance that describes single search expression.</param>
        /// <exception cref="ArgumentNullException"></exception>
        public void AddSearchExpression(ConfigExpression expression)
        {
            if (expression != null)
            {
                Expressions.Add(expression);
            }
            else
            {
                throw new ArgumentNullException($"Null config expression for field {this.Name} was provided.");
            }
        }
        public override string ToString()
        {
            return $"Config field: {Name}; value type: {ValueType}";
        }
        #endregion
    }
}
