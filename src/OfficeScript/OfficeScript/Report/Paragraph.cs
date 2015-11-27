using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = NetOffice.PowerPointApi;

namespace OfficeScript.Report
{
    public class Paragraph
    {
        private PowerPoint.Shape shape;
        private int start;
        private int length;

        public Paragraph(PowerPoint.Shape shape, Dictionary<string, object> parameters)
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
                       return Util.Attr(this, (input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value), Invoke);
                   }),
               remove = (Func<object, Task<object>>)(
                   async (input) =>
                   {
                       this.Remove();
                       return null;
                   })
           };
        }

        /// <summary>
        /// 
        /// </summary>
        private void Remove()
        {
            this.shape.TextFrame.TextRange.Paragraphs(this.start, this.length).Delete();
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
                while (this.shape.TextFrame.TextRange.Paragraphs().Count < this.start - 1)
                {
                    this.shape.TextFrame.TextRange.Paragraphs(this.shape.TextFrame.TextRange.Paragraphs().Count).InsertAfter(Environment.NewLine);
                }
                if (this.shape.TextFrame.TextRange.Paragraphs().Count < this.start)
                {
                    text = Environment.NewLine + text;
                }
                this.shape.TextFrame.TextRange.Paragraphs(this.start, this.length).Text = text;
            }
        }

        public int Count
        {
            get
            {
                return this.shape.TextFrame.TextRange.Paragraphs(this.start,this.length).Count;
            }
        }
        
    }
}
