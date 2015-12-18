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
        }

        private void Init()
        {
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
                       Init();
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
        public int BulletCharacter
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
        
        public String BulletFontName
        {
            get
            {
                return (String)this.format.Bullet.Font.Name;
            }
            set
            {
                this.format.Bullet.Font.Name = value;
            }
        }
        
        public MsoTriState BulletFontBold
        {
            get
            {
                return (MsoTriState)this.format.Bullet.Font.Bold;
            }
            set
            {
                this.format.Bullet.Font.Bold = value;
            }
        }
        
        public int BulletFontSize
        {
            get
            {
                return (int)this.format.Bullet.Font.Size;
            }
            set
            {
                this.format.Bullet.Font.Size = value;
            }
        }
        
        public NetOffice.OfficeApi.ColorFormat BulletFontColor
        {
            get
            {
                return (NetOffice.OfficeApi.ColorFormat)this.format.Bullet.Font.Fill.ForeColor;
            }
            set
            {
                this.format.Bullet.Font.Fill.ForeColor = value;
            }
        }
        
        public float BulletRelativeSize
        {
            get
            {
                return (float)this.format.Bullet.RelativeSize;
            }
            set
            {
                this.format.Bullet.RelativeSize = value;
            }
        }
        
        public MsoTriState BulletVisible
        {
            get
            {
                return (MsoTriState)this.format.Bullet.Visible;
            }
            set
            {
                this.format.Bullet.Visible = value;
            }
        }
        
        public float FirstLineIndent
        {
            get
            {
                return (float)this.format.FirstLineIndent;
            }
            set
            {
                this.format.FirstLineIndent = value;
            }
        }
        
        public float LeftIndent
        {
            get
            {
                return (float)this.format.LeftIndent;
            }
            set
            {
                this.format.LeftIndent = value;
            }
        }
        
        public MsoTriState HangingPunctuation
        {
            get
            {
                return (MsoTriState)this.format.HangingPunctuation;
            }
            set
            {
                this.format.HangingPunctuation = value;
            }
        }
        
        public MsoTriState LineRuleBefore
        {
            get
            {
                return (MsoTriState)this.format.LineRuleBefore;
            }
            set
            {
                this.format.LineRuleBefore = value;
            }
        }
        
        public MsoTriState LineRuleAfter
        {
            get
            {
                return (MsoTriState)this.format.LineRuleAfter;
            }
            set
            {
                this.format.LineRuleAfter = value;
            }
        }
        
        public MsoTriState SpaceBefore
        {
            get
            {
                return (MsoTriState)this.format.LineRuleAfter;
            }
            set
            {
                this.format.LineRuleAfter = value;
            }
        }
        
        public MsoTriState SpaceAfter
        {
            get
            {
                return (MsoTriState)this.format.SpaceAfter;
            }
            set
            {
                this.format.LineRuleAfter = value;
            }
        }
        
        public float SpaceWithin
        {
            get
            {
                return (float)this.format.SpaceWithin;
            }
            set
            {
                this.format.SpaceWithin = value;
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
