using System;
using System.Collections.Generic;
using System.Linq;

namespace SmartOCR
{
    /// <summary>
    /// Describes single search field defined in Excel config file.
    /// </summary>
    internal class ConfigField
    {
        /// <summary>
        /// Data type of searched value.
        /// </summary>
        public string ValueType { get; }

        /// <summary>
        /// Regular expression pattern.
        /// </summary>
        public string RE_Pattern { get; set; }

        /// <summary>
        /// Field name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// String value that is used to check similarity between it and values that match <see cref="RE_Pattern"/>.
        /// </summary>
        public string ExpectedName { get; set; }

        /// <summary>
        /// Collection of search expressions that define search process for current field.
        /// </summary>
        public List<ConfigExpressionBase> Expressions { get; } = new List<ConfigExpressionBase>();

        /// <summary>
        /// Initializes a new <see cref="ConfigField"/> instance with name and value type.
        /// </summary>
        /// <param name="name">Field name.</param>
        /// <param name="value_type">Data type of searched value.</param>
        public ConfigField(string name, string value_type)
        {
            Name = name;
            ValueType = value_type;
        }
        
        /// <summary>
        /// Parses Excel cell contents and gets regular expression pattern and expected field name.
        /// </summary>
        /// <param name="input">String representation of Excel cell contents.</param>
        public void ParseFieldExpression(string input)
        {
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

            RE_Pattern = string.IsNullOrEmpty(splitted_value[0])
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
    }
}
