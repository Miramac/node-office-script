using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.OfficeApi.Enums;

namespace OfficeScript.Report
{
    class Slide
    {
        private PowerPoint.Slide slide;
        private const OfficeScriptType officeScriptType = OfficeScriptType.Slide;
        private PowerPointTags tags;

        public Slide(PowerPoint.Slide slide)
        {
            this.slide = slide;
            this.tags = new PowerPointTags(this.slide);
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
                copy = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.Copy();
                        return new Slide(this.slide).Invoke();
                    }),
                select = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.slide.Select();
                        return new Slide(this.slide).Invoke();
                    }),
                shapes = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        if (input is string)
                        {
                            var tmp = new Dictionary<string, object>();
                            tmp.Add("tag:ctobjectdata.id", input); //remove
                            input = tmp;
                        }
                        input = (input == null) ? new Dictionary<string, object>() : input;
                        return this.Shapes((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                    }),
                addTextbox = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        input = (input == null) ? new Dictionary<string,object>() :  input;
                        return this.AddTextbox((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                    }),
                addPicture = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        input = (input == null) ? new Dictionary<string, object>() : input;
                        return this.AddPicture((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                    }),
                textReplace = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.TextReplace((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                        return this.Invoke();
                    }),
                dispose = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.slide.Dispose();
                        return null;
                    }),
                getType = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return officeScriptType;
                    }
                )
            };
        }

        /// <summary>
        /// Init shape Array
        /// </summary>
        /// <returns></returns>
        private object Shapes(IDictionary<string, object> filter)
        {
            List<object> shapes = new List<object>();

            foreach (PowerPoint.Shape pptShape in this.slide.Shapes)
            {
                var shape = new Shape(pptShape);
                if (shape.TestFilter(filter))
                {
                    shapes.Add(shape.Invoke());
                }
            }

            return shapes.ToArray();
        }


        /// <summary>
        /// Deletes the Slide
        /// </summary>
        private void Remove()
        {
            this.slide.Delete();
        }

        /// <summary>
        /// Duplicate Slide, default position is Slide-Index + 1
        /// </summary>
        private object Duplicate()
        {
            return new Slide(this.slide.Duplicate()[1]).Invoke();
        }

        /// <summary>
        /// Copy Slide
        /// </summary>
        private void Copy()
        {
            this.slide.Copy();
        }
        
        /// <summary>
        /// Not yet Implemented!
        /// </summary>
        private void Sort()
        {
            throw new NotImplementedException("No sorting Algorithm implemented!");
        }

        /// <summary>
        /// AddTextbox and retrun shape object
        /// </summary>
        private object AddTextbox(IDictionary<string, object> parameters)
        {
            object tmpObject;
            float tmpFloat;

            var orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationHorizontal;
            float left = 0;
            float top = 0;
            float height = 100;
            float width = 100;



            //Try to get Shape options: OFFSCRIPT-2
            if (parameters.TryGetValue("left", out tmpObject))
            {
                if (float.TryParse(tmpObject.ToString(), out tmpFloat))
                {
                    left = tmpFloat;
                }
            }
            if (parameters.TryGetValue("top", out tmpObject))
            {
                if (float.TryParse(tmpObject.ToString(), out tmpFloat))
                {
                    top = tmpFloat;
                }
            }
            if (parameters.TryGetValue("height", out tmpObject))
            {
                if (float.TryParse(tmpObject.ToString(), out tmpFloat))
                {
                    height = tmpFloat;
                }
            }
            if (parameters.TryGetValue("width", out tmpObject))
            {
                if (float.TryParse(tmpObject.ToString(), out tmpFloat))
                {
                    width = tmpFloat;
                }
            }

            if (parameters.TryGetValue("texOrientation", out tmpObject))
            {
                switch (tmpObject.ToString().ToLower())
                {
                    case "horizontal":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationHorizontal;
                        break;
                    case "downward":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationDownward;
                        break;
                    case "rotatedfareast":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast;
                        break;
                    case "upward":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationUpward;
                        break;
                    case "vertical":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationVertical;
                        break;
                    case "verticalfareast":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationVerticalFarEast;
                        break;
                    case "mixed": //what is mixed??
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationMixed;
                        break;
                    default:
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationHorizontal;
                        break;
                }
            }

            return new Shape(this.slide.Shapes.AddTextbox(orientation, left, top, width, height)).Invoke();
        }

        /// <summary>
        /// Adds an empty textbox on the given slide
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public object AddPicture(IDictionary<string, object> parameters)
        {
            object tmpObject;
            float tmpFloat;

            float left = 0;
            float top = 0;
            String path = "";

            if (parameters.TryGetValue("left", out tmpObject))
            {
                if (float.TryParse(tmpObject.ToString(), out tmpFloat))
                {
                    left = tmpFloat;
                }
            }
            if (parameters.TryGetValue("top", out tmpObject))
            {
                if (float.TryParse(tmpObject.ToString(), out tmpFloat))
                {
                    top = tmpFloat;
                }
            }
            if (parameters.TryGetValue("path", out tmpObject))
            {
                path = tmpObject.ToString();

                if (!System.IO.File.Exists(path))
                {
                    throw new Exception("Missing file!");
                }
            }
            return new Shape(this.slide.Shapes.AddPicture(path, NetOffice.OfficeApi.Enums.MsoTriState.msoFalse, NetOffice.OfficeApi.Enums.MsoTriState.msoTrue, left, top)).Invoke();
        }

        internal bool TestFilter(IDictionary<string, object> filter)
        {

            //No filter, select all
            if (filter.Keys.Count == 0)
            {
                return true;
            }
            string typeIdentifier;

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
        /// Find and replace in presentation
        /// </summary>
        /// <param name="parameters"></param>
        private void TextReplace(Dictionary<string, object> parameters)
        {
            string find = null;
            string replace = null;
            Dictionary<string, object> replaces = null;
            Dictionary<string, string> newReplaces = null;
            object tmp;

            if (parameters.TryGetValue("find", out tmp))
            {
                find = (string)tmp;
            }
            if (parameters.TryGetValue("replace", out tmp))
            {
                replace = (string)tmp;
            }
            
            if (parameters.TryGetValue("batch", out tmp))
            {
                replaces = (tmp as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value);
                newReplaces = new Dictionary<string, string>();
                // repalce in Array
                foreach(var i in replaces) 
                {
                    var newValue = i.Value.ToString();
                    foreach(var j in replaces) 
                    {
                        newValue = newValue.Replace(j.Key, j.Value.ToString());
                    }
                    newReplaces.Add(i.Key, newValue);
                }
            }

            if(find != null && replace != null){
                TextReplace(find, replace);
            }
            if(newReplaces != null){
                BatchTextReplace(newReplaces);
            }
        }

        /// <summary>
        /// Find and replace in presentation
        /// </summary>
        private void TextReplace(string find, string replace)
        {
            foreach(PowerPoint.Shape shape in slide.Shapes)
            {
                //for textboxes
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    shape.TextFrame.TextRange.Replace(find, replace);
                }
                //for Tables
                else if (shape.HasTable == MsoTriState.msoTrue)
                {
                    foreach (PowerPoint.Row row in shape.Table.Rows)
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
        }

        /// <summary>
        /// Mass Find and replace in slide
        /// </summary>
        private void BatchTextReplace(Dictionary<string, string> replaces)
        {
            foreach(PowerPoint.Shape shape in slide.Shapes)
            {
                if(shape.HasTextFrame == MsoTriState.msoTrue || shape.HasTable == MsoTriState.msoTrue)
                {
                    //for textboxes
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        TextReplace(shape.TextFrame.TextRange, replaces);
                    }
                    //for Tables
                    else if (shape.HasTable == MsoTriState.msoTrue)
                    {
                        foreach (PowerPoint.Row row in shape.Table.Rows)
                        {
                            foreach (PowerPoint.Cell cell in row.Cells)
                            {
                                if (cell.Shape.HasTextFrame == MsoTriState.msoTrue)
                                {
                                    TextReplace(cell.Shape.TextFrame.TextRange, replaces);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void TextReplace(PowerPoint.TextRange textRange, Dictionary<string, string> replaces) 
        {
            var text = textRange.Text;
            foreach (var replace in replaces) 
            {
                if(text.Contains(replace.Key)) 
                {
                    textRange.Replace(replace.Key, replace.Value);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        internal PowerPoint.Slide GetUnderlyingObject()
        {
            return this.slide;
        }

        #region Properties

        public int Pos
        {
            get
            {
                return this.slide.SlideIndex;
            }
            set
            {
                this.slide.MoveTo(value);
            }
        }

        public int Number
        {
            get
            {
                return this.slide.SlideNumber;
            }
        }

        public string Name
        {
            get
            {
                return this.slide.Name;
            }
            set
            {
                this.slide.Name = value;
            }
        }

        #endregion Properties
    }
}
