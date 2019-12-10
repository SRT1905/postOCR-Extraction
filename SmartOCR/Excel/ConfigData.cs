using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartOCR
{
    class ConfigData
    {
        public List<ConfigField> Fields { get; }

        public ConfigField this[int index]
        {
            get
            {
                return Fields[index];
            }
        }

        public ConfigField this[string index]
        {
            get
            {
                return Fields.Where(single_field => single_field.Name == index).First();
            }
        }

        public ConfigData()
        {
            Fields = new List<ConfigField>();
        }

        public ConfigData(params ConfigField[] fields) : this()
        {
            foreach (ConfigField item in fields)
            {
                if (item != null)
                {
                    Fields.Add(item);
                }
            }

        }

        public ConfigData(IEnumerable<ConfigField> fields) : this()
        {
            Fields = fields.ToList();
        }

        public void AddField(ConfigField field)
        {
            if (field != null)
            {
                Fields.Add(field);
            }
            else
            {
                throw new ArgumentNullException("Null config field was provided.");
            }
        }





    }


}
