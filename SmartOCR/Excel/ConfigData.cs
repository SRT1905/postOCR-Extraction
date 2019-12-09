using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartOCR.Excel
{
    class ConfigData
    {
        public List<ConfigField> fields { get; }

        public ConfigField this[int index]
        {
            get
            {
                return fields[index];
            }
        }

        public ConfigField this[string index]
        {
            get
            {
                return fields.Where(single_field => single_field.Name == index).First();
            }
        }

        private ConfigData()
        {
            fields = new List<ConfigField>();
        }

        public ConfigData(params ConfigField[] fields) : this()
        {
            foreach (ConfigField item in fields)
            {
                if (item != null)
                {
                    this.fields.Add(item);
                }
            }
            
        }

        public ConfigData(IEnumerable<ConfigField> fields) : this()
        {
            this.fields = fields.ToList();
        }

        public void AddField(ConfigField field)
        {
            if (field != null)
            {
                this.fields.Add(field);
            }
            else
            {
                throw new ArgumentNullException("Null config field was provided.");
            }
        }





    }


}
