using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = NetOffice.PowerPointApi;
using Office = NetOffice.OfficeApi;

namespace OfficeScript.Report
{
    public class Character
    {
        private PowerPoint.Shape shape;
        private int start;
        private int length;

        public Character(PowerPoint.Shape shape, Dictionary<string, object> parameters)
        {
            this.shape = shape;
            object tmp;
            if (parameters.TryGetValue("start", out tmp))
            {
                this.start = (int)tmp;
            }
            else
            {
                this.start = -1;
            }
            if (parameters.TryGetValue("length", out tmp))
            {
                this.length = (int)tmp;
            }
            else
            {
                this.length = -1;
            }

           
        }

        /// <summary>
        /// Retuns an object with async functions for node.js
        /// </summary>
        /// <returns></returns>
        public object Invoke()
        {
            return new
           {
               attr = (Func<object, Task<object>>)(
                   async (input) =>
                   {
                       if (input is string)
                       {
                           var tmp = new Dictionary<string, object>();
                           tmp.Add("name", input);
                           input = tmp;
                       }
                       return Util.Attr(this, (input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value), Invoke);
                   }),
               remove = (Func<object, Task<object>>)(
                   async (input) =>
                   {
                       this.Remove();
                       return null;
                   }),
                format = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return new Format(this).Invoke();
                    }),
                font = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return new Font(this, this.Invoke).Invoke();
                    })
            };
        }

        /// <summary>
        /// 
        /// </summary>
        private void Remove()
        {
            this.shape.TextFrame.TextRange.Characters(this.start, this.length).Delete();
        }

        /// <summary>
        /// Get or Set the Text-Property for this element.
        /// </summary>
        public string Text
        {
            get
            {
                return this.shape.TextFrame.TextRange.Paragraphs(this.start, this.length).Text.TrimEnd();
            }
            set
            {
                string text = value;
                int i;

                for (i = this.shape.TextFrame.TextRange.Paragraphs().Count; i < this.start; i++)
                {
                    this.shape.TextFrame.TextRange.Paragraphs(this.shape.TextFrame.TextRange.Paragraphs().Count).InsertAfter(Environment.NewLine);
                }

                //if (this.shape.TextFrame.TextRange.Paragraphs().Count < this.start)
                //{
                    //text = Environment.NewLine + text;
                //}
                this.shape.TextFrame.TextRange.Paragraphs(this.start, this.length).Text = text;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public int Count
        {
            get
            {
                return this.shape.TextFrame.TextRange.Paragraphs(this.start, this.length).Lines().Count;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public NetOffice.OfficeApi.TextRange2 GetUnderlyingObject()
        {
            return this.shape.TextFrame2.TextRange.Paragraphs(this.start, this.length);
        }

    }
}
