﻿namespace SmartOCR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Describes single search field defined in Excel config file.
    /// </summary>
    public class ConfigField // TODO: implement Soundex usage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigField"/> class.
        /// Instance is initialized with name and value type.
        /// </summary>
        /// <param name="name">Field name.</param>
        /// <param name="valueType">Data type of searched value.</param>
        public ConfigField(string name, string valueType)
        {
            this.Name = name;
            this.ValueType = valueType;
        }

        /// <summary>
        /// Gets collection of search expressions that define search process for current field.
        /// </summary>
        public List<ConfigExpression> Expressions { get; } = new List<ConfigExpression>();

        /// <summary>
        /// Gets or sets string value that is used to check similarity between it and values that match <see cref="RegExPattern"/>.
        /// </summary>
        public string ExpectedName { get; set; }

        /// <summary>
        /// Gets or sets regular expression pattern.
        /// </summary>
        public string RegExPattern { get; set; }

        /// <summary>
        /// Gets field name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets data type of searched value.
        /// </summary>
        public string ValueType { get; }

        /// <summary>
        /// Gets a value indicating whether field identifier should be searched using <see cref="ParagraphContainer.Soundex"/> instead of <see cref="ParagraphContainer.Text"/>.
        /// </summary>
        public bool IsSoundex { get; private set; }

        /// <summary>
        /// Parses Excel cell contents and gets regular expression pattern and expected field name.
        /// </summary>
        /// <param name="input">String representation of Excel cell contents.</param>
        public void ParseFieldExpression(string input)
        {
            if (input == null)
            {
                this.RegExPattern = this.Name;
                this.ExpectedName = this.Name;
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

            this.RegExPattern = string.IsNullOrEmpty(splittedValue[0])
                ? this.Name
                : splittedValue[0];
            this.ExpectedName = string.IsNullOrEmpty(splittedValue[1])
                ? this.Name
                : splittedValue[1];
        }

        /// <summary>
        /// Inserts <see cref="ConfigExpression"/> instance to expression collection.
        /// </summary>
        /// <param name="expression"><see cref="ConfigExpression"/> instance that describes single search expression.</param>
        /// <exception cref="ArgumentNullException">Empty config expression.</exception>
        public void AddSearchExpression(ConfigExpression expression)
        {
            if (expression != null)
            {
                this.Expressions.Add(expression);
            }
            else
            {
                throw new ArgumentNullException($"Null config expression for field {this.Name} was provided.");
            }
        }

        /// <inheritdoc/>
        public override string ToString()
        {
            return $"Config field: {this.Name}; value type: {this.ValueType}";
        }
    }
}
