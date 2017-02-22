using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using NetOffice.OfficeApi.Enums;
using PowerPoint = NetOffice.PowerPointApi;

namespace OfficeScript.Report
{
    class Shape
    {
        private PowerPoint.Shape shape;
        private const OfficeScriptType officeScriptType = OfficeScriptType.Shape;
        private PowerPointTags tags;

        public Shape(PowerPoint.Shape shape)
        {
            this.shape = shape;
            this.tags = new PowerPointTags(this.shape);
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
                tags = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.tags.Invoke();
                    }),
                remove = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.Remove();
                        return null;
                    }),
                duplicate = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.Duplicate();
                    }),
                paragraph = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        input = (input == null) ? new Dictionary<string, object>() : input;
                        return new Paragraph(this.shape, (input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value)).Invoke();
                    }),
                character = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        input = (input == null) ? new Dictionary<string, object>() : input;
                        return new Character(this.shape, (input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value)).Invoke();
                    }),
                textReplace = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.TextReplace((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                        return this.Invoke();
                    }),
                exportAs = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        if (input is string)
                        {
                            var tmp = new Dictionary<string, object>();
                            tmp.Add("path", input);
                            input = tmp;
                        }
                        ExportAs((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                        return null;
                    }),
                hasObject = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.HasObject((string)input);
                    }),
                getZindex = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.GetZindex();
                    }),
                setZindex = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.SetZindex((string)input);
                        return this.Invoke();
                    }),
                dispose = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.shape.Dispose();
                        return null;
                    }),
                getType = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return officeScriptType;
                    })

            };
        }

        /// <summary>
        /// Duplicate this Shape
        /// </summary>
        /// <returns>Shape</returns>
        private object Duplicate()
        {
             return new Shape(this.shape.Duplicate()[1]).Invoke();
        }
        
        /// <summary>
        /// Deletes the Shape
        /// </summary>
        private void Remove()
        {
            this.shape.Delete();
            this.shape.Dispose();
        }

        /// <summary>
        /// Search and replace
        /// </summary>
        /// <param name="parameters"></param>
        private void TextReplace(Dictionary<string, object> parameters)
        {
            string find = null;
            string replace = null;
            object tmp;


            if (parameters.TryGetValue("find", out tmp))
            {
                find = (string)tmp;
            }
            if (parameters.TryGetValue("replace", out tmp))
            {
                replace = (string)tmp;
            }
     
            if(find != null && replace != null){
                TextReplace(find, replace);
            }
        }

        /// <summary>
        /// Use PPT buildin search and replace function
        /// </summary>
        private void TextReplace(string find, string replace)
        {
            //for textboxes
            if (this.shape.HasTextFrame == MsoTriState.msoTrue)
            {
                this.shape.TextFrame.TextRange.Replace(find, replace);
            }
            //for Tables
            else if (this.shape.HasTable == MsoTriState.msoTrue)
            {
                foreach (PowerPoint.Row row in this.shape.Table.Rows)
                {
                    foreach (PowerPoint.Cell cell in row.Cells)
                    {
                        if (cell.Shape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            cell.Shape.TextFrame.TextRange.Replace(find, replace);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Export shape as picture
        /// </summary>
        /// <param name="path"></param>
        private void ExportAs(IDictionary<string, object> parameters)
        {
            string path = (string)parameters["path"];
            string type = "png";
           // float heigth = 542;
           // float width = 722;
           // float scale = 2;
            float heigth = this.shape.Height;
            float width =  this.shape.Width;
            float scale =  1;

            object tmp;

            PowerPoint.Enums.PpShapeFormat ppShapeFormat = PowerPoint.Enums.PpShapeFormat.ppShapeFormatPNG;

            if (parameters.TryGetValue("type", out tmp))
            {
                type = (string)tmp;
            }
            if (parameters.TryGetValue("heigth", out tmp))
            {
                heigth = (float)tmp;
            }
            if (parameters.TryGetValue("width", out tmp))
            {
                width = (float)tmp;
            }
            if (parameters.TryGetValue("scale", out tmp))
            {
                scale = (float)tmp;
            }
            //couse
            switch (type.ToLower())
            {
                case "png":
                    ppShapeFormat = PowerPoint.Enums.PpShapeFormat.ppShapeFormatPNG;
                    break;
                case "wmf":
                    ppShapeFormat = PowerPoint.Enums.PpShapeFormat.ppShapeFormatWMF;
                    break;
                case "bmp":
                    ppShapeFormat = PowerPoint.Enums.PpShapeFormat.ppShapeFormatBMP;
                    break;
                case "gif":
                    ppShapeFormat = PowerPoint.Enums.PpShapeFormat.ppShapeFormatGIF;
                    break;
                case "jpg":
                    ppShapeFormat = PowerPoint.Enums.PpShapeFormat.ppShapeFormatJPG;
                    break;
                case "emf":
                    ppShapeFormat = PowerPoint.Enums.PpShapeFormat.ppShapeFormatEMF;
                    break;
                default:
                    ppShapeFormat = PowerPoint.Enums.PpShapeFormat.ppShapeFormatPNG;
                    break;
            }
            this.shape.Export(path, ppShapeFormat, width * scale, heigth * scale, PowerPoint.Enums.PpExportMode.ppRelativeToSlide);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private bool HasObject(string name)
        {
            switch(name.ToLower())
            {
                case "chart":
                    return (this.shape.HasChart == MsoTriState.msoTrue);
                case "diagram":
                    return (this.shape.HasDiagram == MsoTriState.msoTrue);
                case "table":
                    return (this.shape.HasTable == MsoTriState.msoTrue);
                case "img":
                case "image":
                case "picture":
                    return (this.shape.Type == MsoShapeType.msoPicture);
                case "text":
                case "textframe":
                    return (this.shape.HasTextFrame == MsoTriState.msoTrue);
                default:
                    return false;
            }
        }

        /// <summary>
        /// Test if the given Filter matches the shape. Filter can be Tags or Properties. 
        /// eg: {"tag:mytag", "Some Value", "attr:Name", "Shape Name"}
        /// </summary>
        /// <param name="filter"></param>
        /// <returns></returns>
        internal bool TestFilter(IDictionary<string, object> filter)
        {

            //No filter, select all
            if (filter.Keys.Count == 0)
            {
                return true;
            }
            string typeIdentifier;

            //Test Tag selectors
            typeIdentifier = "tag:";
            foreach (string key in filter.Keys.Where(w => w.StartsWith(typeIdentifier)).ToArray())
            {
                string tagName = key.Substring(typeIdentifier.Length);
                string tagValue = this.tags.Get(tagName);
                string[] values = filter[key].ToString().Split(',');
                for (int i = 0; i < values.Length; i++)
                {
                    string val = values[i].Trim();
                    if (tagValue == val)
                    {
                        return true;
                    }
                }
            }

            //Test Attr selectors
            typeIdentifier = "attr:";
            foreach (string key in filter.Keys.Where(w => w.StartsWith(typeIdentifier)).ToArray())
            {
                string attrName = key.Substring(typeIdentifier.Length);
                string attrValue = this.GetType().GetProperty(attrName).GetValue(this, null).ToString();
                string[] values = filter[key].ToString().Split(',');
                for (int i = 0; i < values.Length; i++)
                {
                    string val = values[i].Trim();
                    if (attrValue == val)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        internal PowerPoint.Shape GetUnderlyingObject()
        {
            return this.shape;
        }

        /// <summary>
        /// 
        /// </summary>
        internal PowerPointTags GetTags()
        {
            return new PowerPointTags(this.shape);
        }

        public void SetZindex(string order)
        {
            switch (order.ToLower())
            {
                case "forward":
                    this.shape.ZOrder(MsoZOrderCmd.msoBringForward);
                    break;
                case "backward":
                    this.shape.ZOrder(MsoZOrderCmd.msoSendBackward);
                    break;
                case "front":
                    this.shape.ZOrder(MsoZOrderCmd.msoBringToFront);
                    break;
                case "back":
                    this.shape.ZOrder(MsoZOrderCmd.msoSendBackward);
                    break;
                case "beforetext":
                    this.shape.ZOrder(MsoZOrderCmd.msoBringInFrontOfText);
                    break;
                case "behindtext":
                    this.shape.ZOrder(MsoZOrderCmd.msoSendBehindText);
                    break;
            }
        }

        public int GetZindex()
        {
            return this.shape.ZOrderPosition;    
        }

        #region Properties

        public string Name
        {
            get
            {
                return this.shape.Name;
            }
            set
            {
                this.shape.Name = value;
            }
        }
        public string Text
        {
            get
            {
                if(this.shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    return this.shape.TextFrame.TextRange.Text;
                } else
                {
                    return null;
                }
            }
            set
            {
                if (this.shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    this.shape.TextFrame.TextRange.Text = value;
                }               
            }
        }

        /// <summary>
        /// Get or Set the Top-Property for this element.
        /// </summary>
        public double Top
        {
            get
            {
                return this.shape.Top;
            }
            set
            {
                this.shape.Top = (float)value;
            }
        }

        /// <summary>
        /// Get or Set the Left-Property for this element.
        /// </summary>
        public double Left
        {
            get
            {
                return this.shape.Left;
            }
            set
            {
                this.shape.Left = (float)value;
            }
        }
        /// <summary>
        /// Get or Set the Height-Property for this element.
        /// </summary>
        public double Height
        {
            get
            {
                return this.shape.Height;
            }
            set
            {
                this.shape.Height = (float)value;
            }
        }
        /// <summary>
        /// Get or Set the Width-Property for this element.
        /// </summary>
        public double Width
        {
            get
            {
                return this.shape.Width;
            }
            set
            {
                this.shape.Width = (float)value;
            }
        }
        /// <summary>
        /// Get or Set the Rotation-Property for this element.
        /// </summary>
        public double Rotation
        {
            get
            {
                return this.shape.Rotation;
            }
            set
            {
                this.shape.Rotation = (float)value;
            }
        }

        /// <summary>
        /// Get or Set the Fill-Property for this element.
        /// </summary>
        public string Fill
        {
            get
            {
                string bgr = "#" + this.shape.Fill.ForeColor.RGB.ToString("x6");
                return Util.BGRtoRGB(bgr);
            }
            set
            {
                this.shape.Fill.ForeColor.RGB = ColorTranslator.FromHtml(Util.BGRtoRGB(value)).ToArgb();
            }
        }

        /// <summary>
        /// Get or Set the Alt-Text for this element.
        /// </summary>
        public string AltText
        {
            get
            {
                return this.shape.AlternativeText;
            }
            set
            {
                this.shape.AlternativeText = value;
            }
        }

        /// <summary>
        /// Get or Set the Alt-Text for this element.
        /// </summary>
        public string Title
        {
            get
            {
                return this.shape.Title;
            }
            set
            {
                this.shape.Title = value;
            }
        }

        /// <summary>
        /// Get or Set the Top-Property for this element.
        /// </summary>
        public string Type
        {
            get
            {
                return this.shape.Type.ToString();
            }set
            {
                // do nothing :)
            }
        }

        // public Dictionary<string, object> Table
        public object[] Table
        {
            get
            {
                object[] returnTable = null;
                int rowCount, columnCount;
                if (this.shape.HasTable == MsoTriState.msoTrue)
                {
                    returnTable = new object[shape.Table.Rows.Count];

                    rowCount = 0;
                    foreach (PowerPoint.Row row in shape.Table.Rows)
                    {
                        object[] cells = new object[shape.Table.Columns.Count];
                        columnCount = 0;
                        foreach (PowerPoint.Cell cell in row.Cells)
                        {
                            cells[columnCount++] = new Shape(cell.Shape).Invoke();
                            // cells[columnCount++] = (cell.Shape.HasTextFrame == MsoTriState.msoTrue) ? cell.Shape.TextFrame.TextRange.Text : null;
                        }
                        returnTable[rowCount++] = cells;

                    }
                }
                return returnTable;
            }
        }


        #endregion Properties
    }
}
