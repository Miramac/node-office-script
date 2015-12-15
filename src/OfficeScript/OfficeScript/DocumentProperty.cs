using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = NetOffice.PowerPointApi;

namespace OfficeScript.Report
{
    class DocumentProperty
    {
        private dynamic element;

        public DocumentProperty(PowerPoint.Presentation presentation)
        {
            this.element = presentation;
        }

        public object Invoke()
        {
            return new
            {
                getCustomProperty = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.GetCustomProperty((string)input);
                    }
                ),
                getBuiltinProperty = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.GetBuiltinProperty((string)input);
                    }
                ),
                setBuiltinProperty = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.SetBuiltinProperty((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                        return this.Invoke();
                    }
                ),
                setCustomProperty = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.SetCustomProperty((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                        return this.Invoke();
                    }
                ),
                getAllCustomProperties = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.GetAllCustomProperties();
                    }
                ),
                getAllBuiltinProperties = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.GetAllBuiltinProperties();
                    }
                )
            };
        }
        private Dictionary<string, string> GetAllCustomProperties()
        {
            NetOffice.OfficeApi.DocumentProperties properties;
            properties = (NetOffice.OfficeApi.DocumentProperties)this.element.CustomDocumentProperties;
            var props = new Dictionary<string, string>();

            foreach (NetOffice.OfficeApi.DocumentProperty prop in properties)
            {
                try
                {
                    if (prop.Value != null)
                    {
                        props.Add(prop.Name.ToString(), prop.Value.ToString());
                    }
                }
                catch
                {

                }
            }
            return props;
        }

        private Dictionary<string, string> GetAllBuiltinProperties()
        {
            NetOffice.OfficeApi.DocumentProperties properties;
            properties = (NetOffice.OfficeApi.DocumentProperties)this.element.BuiltInDocumentProperties;
            var props = new Dictionary<string, string>();

            foreach (NetOffice.OfficeApi.DocumentProperty prop in properties)
            {
                try
                {
                    if (prop.Value != null)
                    {
                        props.Add(prop.Name.ToString(), prop.Value.ToString());
                    }
                }
                catch
                {

                }
            }
            return props;
        }

        private string GetCustomProperty(string propertyName)
        {
            NetOffice.OfficeApi.DocumentProperties properties;
            properties = (NetOffice.OfficeApi.DocumentProperties)this.element.CustomDocumentProperties;

            foreach (NetOffice.OfficeApi.DocumentProperty prop in properties)
            {
                if (prop.Name.ToString().ToUpper() == propertyName.ToUpper())
                {
                    return prop.Value.ToString();
                }
            }
            return null;
        }

        private void SetCustomProperty(Dictionary<string, object> parameters)
        {
            NetOffice.OfficeApi.DocumentProperties properties;
            properties = (NetOffice.OfficeApi.DocumentProperties)this.element.CustomDocumentProperties;
            String propertyName = (string)parameters["prop"];
            String propertyVal = (string)parameters["value"];

            if (GetCustomProperty(propertyName) != null)
            {
                properties[propertyName].Delete();
            }

            properties.Add(propertyName, false, NetOffice.OfficeApi.Enums.MsoDocProperties.msoPropertyTypeString, propertyVal);
        }

        private string GetBuiltinProperty(string propertyName)
        {
            NetOffice.OfficeApi.DocumentProperties properties;
            properties = (NetOffice.OfficeApi.DocumentProperties)this.element.BuiltInDocumentProperties;

                foreach (NetOffice.OfficeApi.DocumentProperty prop in properties)
                {
                    if (prop.Name.ToString() == propertyName)
                    {
                        return prop.Value.ToString();
                    }
                }
            return null;
        }

        private void SetBuiltinProperty(Dictionary<string, object> parameters)
        {
            NetOffice.OfficeApi.DocumentProperties properties;
            properties = (NetOffice.OfficeApi.DocumentProperties)this.element.BuiltInDocumentProperties;
            String propertyName = (string)parameters["prop"];
            String propertyVal = (string)parameters["value"];

            if (GetBuiltinProperty(propertyName) != null)
            {
                properties[propertyName].Value = propertyVal;
    
            }
        }

    }
}
