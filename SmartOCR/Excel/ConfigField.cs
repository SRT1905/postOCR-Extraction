using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartOCR.Excel
{
    class ConfigField
    {
        public string ValueType { get; }
        public string RE_Pattern { get; }
        public string Name { get; }
        public string ExpectedName { get; }
        public List<ConfigExpression> Expressions { get; }

        private ConfigField()
        {
            Expressions = new List<ConfigExpression>();
        }

        public ConfigField(string name, string value_type) : this()
        {
            this.Name = name;
            this.ValueType = value_type;
        }

        public void ParseFieldExpression(string input)
        {
            
        }

        public void AddSearchExpression(ConfigExpression expression)
        {
            this.Expressions.Add(expression);
        }

    }
}
