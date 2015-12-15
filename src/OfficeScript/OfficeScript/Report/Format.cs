using System;
using NetOffice.OfficeApi.Enums;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;

namespace OfficeScript.Report
{

    class Format
    {
        private NetOffice.OfficeApi.ParagraphFormat2 format;
        private Paragraph paragraph;

        public Format(NetOffice.OfficeApi.ParagraphFormat2 format)
        {
            this.format = format;
        }

        public Format(Paragraph paragraph)
        {
            this.paragraph = paragraph;
            this.format = this.paragraph.GetUnderlyingObject().ParagraphFormat;
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
                       return Util.Attr(this, (input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value), paragraph.Invoke);
                   }),
            };
      }

        /// <summary>
        /// Get or Set the Alignment-Property for this element.
        /// Parameters are: "left", "right", "center"
        /// </summary>
        public string Alignment
        {
            get
            {
                switch (this.format.Alignment)
                {
                    case MsoParagraphAlignment.msoAlignLeft:
                        return "left";
                    case MsoParagraphAlignment.msoAlignRight:
                        return "right";
                    case MsoParagraphAlignment.msoAlignCenter:
                        return "center";
                    default:
                        return this.format.Alignment.ToString();
                }
            }
            set
            {
                switch (value.ToLower())
                {
                    case "left":
                        this.format.Alignment = MsoParagraphAlignment.msoAlignLeft;
                        break;
                    case "right":
                        this.format.Alignment = MsoParagraphAlignment.msoAlignRight;
                        break;
                    case "center":
                        this.format.Alignment = MsoParagraphAlignment.msoAlignCenter;
                        break;
                }
            }
        }

        /// <summary>
        /// Get or Set the Bullet-Property for this element.
        /// </summary>
        public int Bullet
        {
            get
            {
                return (int)this.format.Bullet.Character;
            }
            set
            {
                this.format.Bullet.Character = value;
            }
        }

        /// <summary>
        /// Get or Set the Indent-Property for this element.
        /// </summary>
        public int IndentLevel
        {
            get
            {
                return this.format.IndentLevel;
            }
            set
            {
                this.format.IndentLevel = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        internal NetOffice.OfficeApi.ParagraphFormat2 GetUnderlyingObject()
        {
            return this.format;
        }


        //http://codereview/#3WF
        public void Copy(Format src)
        {
            NetOffice.OfficeApi.ParagraphFormat2 srcFormat = src.GetUnderlyingObject();

            //Bullets
            this.format.Bullet.Font.Name = srcFormat.Bullet.Font.Name;
            this.format.Bullet.Font.Bold = srcFormat.Bullet.Font.Bold;
            this.format.Bullet.Font.Size = srcFormat.Bullet.Font.Size;
            this.format.Bullet.Font.Fill.ForeColor = srcFormat.Bullet.Font.Fill.ForeColor;
            this.format.Bullet.Character = srcFormat.Bullet.Character;
            this.format.Bullet.RelativeSize = srcFormat.Bullet.RelativeSize;
            this.format.Bullet.Visible = srcFormat.Bullet.Visible;
            //Indent
            this.format.FirstLineIndent = srcFormat.FirstLineIndent;
            this.format.IndentLevel = srcFormat.IndentLevel;
            this.format.LeftIndent = srcFormat.LeftIndent;
            this.format.HangingPunctuation = srcFormat.HangingPunctuation;
            this.format.LineRuleBefore = srcFormat.LineRuleBefore;
            this.format.LineRuleAfter = srcFormat.LineRuleAfter;
            //Spacing
            this.format.SpaceBefore = srcFormat.SpaceBefore;
            this.format.SpaceAfter = srcFormat.SpaceAfter;
            this.format.SpaceWithin = srcFormat.SpaceWithin;
        }
    }
}
