using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartOCR
{
    class ConfigField
    {
        public string ValueType { get; }
        public string RE_Pattern { get; set; }
        public string Name { get; }
        public string ExpectedName { get; set; }
        public List<ConfigExpression> Expressions { get; }

        private ConfigField()
        {
            Expressions = new List<ConfigExpression>();
        }

        public ConfigField(string name, string value_type) : this()
        {
            Name = name;
            ValueType = value_type;
        }
        
        public void ParseFieldExpression(string input)
        {
            string[] splitted_value = input.Split(';');
            RE_Pattern = string.IsNullOrEmpty(splitted_value[0])
                ? Name
                : splitted_value[0];
            ExpectedName = string.IsNullOrEmpty(splitted_value[1])
                ? Name
                : splitted_value[1];
        }

        public void SetFieldExpression(string pattern, string expected_name)
        {
            RE_Pattern = pattern;
            ExpectedName = expected_name;
        }

        public void AddSearchExpression(ConfigExpression expression)
        {
            Expressions.Add(expression);
        }

    }
}
