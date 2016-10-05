using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using NetOffice.OfficeApi.Enums;
using Office = NetOffice.OfficeApi;

namespace OfficeScript.Report
{
    class Font
    {
        private Office.Font2 font;
        private Paragraph paragraph = null;
        private Character character = null;
        private bool disposed;
        private Func<object> parentInvoke;

        public Font(Office.Font2 font, Func<object> parentInvoke)
        {
            this.font = font;
            this.parentInvoke = parentInvoke;
        }

        public Font(Paragraph paragraph, Func<object> parentInvoke)
        {
            this.paragraph = paragraph;
            this.parentInvoke = parentInvoke;
        }

        public Font(Character character, Func<object> parentInvoke)
        {
            this.character = character;
            this.parentInvoke = parentInvoke;
        }

        private void Init()
        {
            if(this.paragraph != null)
            {
                this.font = this.paragraph.GetUnderlyingObject().Font;
            }
            if (this.character != null)
            {
                this.font = this.character.GetUnderlyingObject().Font;
            }
        }

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
                       Init();
                       return Util.Attr(this, (input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value), parentInvoke);
                   }),
            };
        }

        /// <summary>
        /// Get or Set the Bold-Property for this element.
        /// </summary>
        public bool Bold
        {
            get
            {
                return (this.font.Bold == MsoTriState.msoTrue ? true : false);
            }
            set
            {
                if (value == true)
                {
                    this.font.Bold = MsoTriState.msoTrue;
                }
                else
                {
                    this.font.Bold = MsoTriState.msoFalse;
                }
            }
        }
        /// <summary>
        /// Get or Set the Italic-Property for this element.
        /// </summary>
        public bool Italic
        {
            get
            {
                return (this.font.Italic == MsoTriState.msoTrue ? true : false);
            }
            set
            {
                if (value == true)
                {
                    this.font.Italic = MsoTriState.msoTrue;
                }
                else
                {
                    this.font.Italic = MsoTriState.msoFalse;
                }
            }
        }
        /// <summary>
        /// Get or Set the Color-Property for this element.
        /// </summary>
        public string Color
        {
            get
            {
                string bgr = "#" + this.font.Fill.ForeColor.RGB.ToString("x6");
                return Util.BGRtoRGB(bgr);
            }
            set
            {
                this.font.Fill.ForeColor.RGB = ColorTranslator.FromHtml(Util.BGRtoRGB(value)).ToArgb();
            }
        }

        /// <summary>
        /// Get or Set the Size-Property for this element.
        /// </summary>
        public float Size
        {
            get
            {
                return this.font.Size;
            }
            set
            {
                this.font.Size = value;
            }
        }
        /// <summary>
        /// Get or Set the Name-Property for this element.
        /// </summary>
        public string Name
        {
            get
            {
                return this.font.Name;
            }
            set
            {
                this.font.Name = value;
            }
        }

        public void Copy(Font src)
        {
            this.Bold = src.Bold;
            this.Italic = src.Italic;
            this.Name = src.Name;
            this.Size = src.Size;
        }


        internal NetOffice.OfficeApi.Font2 GetUnderlyingObject()
        {
            return this.font;
        }

    }
}
