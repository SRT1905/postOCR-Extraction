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
        public List<ConfigExpressionBase> Expressions { get; } = new List<ConfigExpressionBase>();
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
            var splitted_value = input.Split(';').ToList();

            while (splitted_value.Count != 2)
            {
                splitted_value[0] = $"{splitted_value[0]};{splitted_value[1]}";
                for (int i = 2; i < splitted_value.Count; i++)
                {
                    splitted_value[i - 1] = splitted_value[i];
                }
                splitted_value.RemoveAt(splitted_value.Count - 1);
            }

            RegExPattern = string.IsNullOrEmpty(splitted_value[0])
                ? Name
                : splitted_value[0];
            ExpectedName = string.IsNullOrEmpty(splitted_value[1])
                ? Name
                : splitted_value[1];
        }
        /// <summary>
        /// Inserts <see cref="ConfigExpression"/> instance to expression collection.
        /// </summary>
        /// <param name="expression"><see cref="ConfigExpression"/> instance that describes single search expression.</param>
        /// <exception cref="ArgumentNullException"></exception>
        public void AddSearchExpression(ConfigExpressionBase expression)
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
