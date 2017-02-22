using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.OfficeApi.Enums;


namespace OfficeScript.Report
{
    public class Presentation : IDisposable
    {
        private bool disposed;
        private PowerPoint.Presentation presentation;
        private const OfficeScriptType officeScriptType = OfficeScriptType.Presentation;
        private bool closePresentation = true;
        private PowerPointTags tags;
        private DocumentProperty properties;
        public Presentation(PowerPoint.Presentation presentation)
        {
            this.presentation = presentation;
            this.tags = new PowerPointTags(this.presentation);
            this.properties = new DocumentProperty(this.presentation);
        }

        // Destruktor
        ~Presentation()
        {
            Dispose(false);
        }

        #region Dispose

        // Implement IDisposable.
        // Do not make this method virtual.
        // A derived class should not be able to override this method.
        public void Dispose()
        {
            Dispose(true);
            // This object will be cleaned up by the Dispose method.
            // Therefore, you should call GC.SupressFinalize to
            // take this object off the finalization queue
            // and prevent finalization code for this object
            // from executing a second time.
            GC.SuppressFinalize(this);
        }
        // Dispose(bool disposing) executes in two distinct scenarios.
        // If disposing equals true, the method has been called directly
        // or indirectly by a user's code. Managed and unmanaged resources
        // can be disposed.
        // If disposing equals false, the method has been called by the
        // runtime from inside the finalizer and you should not reference
        // other objects. Only unmanaged resources can be disposed.
        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed
                // and unmanaged resources.
                if (disposing)
                {
                    if (this.closePresentation)
                    {
                        this.presentation.Saved = MsoTriState.msoTrue;
                        this.presentation.Close();
                    }
                    this.presentation.Dispose();

                }

                // Note disposing has been done.
                this.disposed = true;

            }
        }
        #endregion Dispose

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
                properties = (Func<object, Task<object>>)(
                async (input) =>
                {
                    return this.properties.Invoke();
                }),
                save = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.Save();
                        return null;
                    }
                ),
                saveAs = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        if (input is string)
                        {
                            var tmp = new Dictionary<string, object>();
                            tmp.Add("name", input);
                            input = tmp;
                        }
                        this.SaveAs((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                        return null;
                    }
                ),
                saveAsCopy = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        if (input is string)
                        {
                            var tmp = new Dictionary<string, object>();
                            tmp.Add("name", input);
                            input = tmp;
                        }
                        this.SaveAsCopy((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                        return null;
                    }
                ),
                close = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.Dispose();
                        return null;
                    }
                ),
                slides = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        input = (input == null) ? new Dictionary<string, object>() : input;
                        return this.Slides((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                    }
                ),
                addSlide = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        if (input is string)
                        {
                            var tmp = new Dictionary<string, object>();
                            tmp.Add("name", input);
                            input = tmp;
                        }
                        if (input is int)
                        {
                            var tmp = new Dictionary<string, object>();
                            tmp.Add("pos", input);
                            input = tmp;
                        }
                        input = (input == null) ? new Dictionary<string, object>() : input;
                        return this.AddSlide((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                    }
                ),
                textReplace = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.TextReplace((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                        return this.Invoke();
                    }
                ),
                getSelectedShape = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.GetSelectedShape();
                    }
                ),
                getActiveSlide = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.GetActiveSlide();
                    }
                ),
                pasteSlide = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        
                        return this.PasteSlide((int)input);
                    }),
                dispose = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.presentation.Dispose();
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


        #region Save
        private void Save()
        {
            this.presentation.Save();
        }

        private void SaveAs(IDictionary<string, object> parameters)
        {
            string name = (string)(parameters as IDictionary<string, object>)["name"];
            string type = "pptx";
            object tmp;
            if (parameters.TryGetValue("type", out tmp))
            {
                type = (string)tmp;
            }
            this.SaveAs(name, type);
        }


        private void SaveAs(string fileName, string fileType)
        {
            this.SaveAs(fileName, fileType, false);
        }

        private void SaveAsCopy(IDictionary<string, object> parameters)
        {
            string name = (string)(parameters as IDictionary<string, object>)["name"];
            string type = "pptx";
            object tmp;
            if (parameters.TryGetValue("type", out tmp))
            {
                type = (string)tmp;
            }
            this.SaveAs(name, type, true);
        }

        private void SaveAsCopy(string fileName, string fileType)
        {
            this.SaveAs(fileName, fileType, true);
        }

        private void SaveAs(string fileName, string fileType, bool isCopy)
        {
            PowerPoint.Enums.PpSaveAsFileType pptFileType;
            switch (fileType.ToLower())
            {
                case "pdf":
                    pptFileType = PowerPoint.Enums.PpSaveAsFileType.ppSaveAsPDF;
                    break;
                case "pptx":
                    pptFileType = PowerPoint.Enums.PpSaveAsFileType.ppSaveAsOpenXMLPresentation;
                    break;
                case "ppt":
                    pptFileType = PowerPoint.Enums.PpSaveAsFileType.ppSaveAsPresentation;
                    break;
                default:
                    pptFileType = PowerPoint.Enums.PpSaveAsFileType.ppSaveAsOpenXMLPresentation;
                    break;
            }
            if (isCopy)
            {
                this.presentation.SaveCopyAs(fileName, pptFileType);
            }
            else
            {
                this.presentation.SaveAs(fileName, pptFileType);
            }
        }
        #endregion save

        /// <summary>
        /// Add a new Slide
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        private object AddSlide(Dictionary<string, object> parameters)
        {
            var pos = this.presentation.Slides.Count + 1;
            var layout = NetOffice.PowerPointApi.Enums.PpSlideLayout.ppLayoutBlank;
            object tmpObject;
            int tmpInt;

            if (parameters.TryGetValue("pos", out tmpObject))
            {
                if (int.TryParse(tmpObject.ToString(), out tmpInt))
                {
                    pos = tmpInt;
                }
            }
            if (parameters.TryGetValue("layout", out tmpObject))
            {
                switch(tmpObject.ToString().ToLower())
                {
                    case "blank":
                            layout = NetOffice.PowerPointApi.Enums.PpSlideLayout.ppLayoutBlank;
                            break;
                    case "chart":
                        layout = NetOffice.PowerPointApi.Enums.PpSlideLayout.ppLayoutChart;
                        break;
                    case "text":
                        layout = NetOffice.PowerPointApi.Enums.PpSlideLayout.ppLayoutText;
                        break;
                    case "chartandtext":
                        layout = NetOffice.PowerPointApi.Enums.PpSlideLayout.ppLayoutChartAndText;
                        break;
                    case "custom":
                        layout = NetOffice.PowerPointApi.Enums.PpSlideLayout.ppLayoutCustom;
                        break;
                }
            }
            
            return new Slide(this.presentation.Slides.Add(pos, layout)).Invoke();
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
                BatchTextReplace(new Dictionary<string, string>() {{find, replace }});
            }
            if(newReplaces != null){
                BatchTextReplace(newReplaces);
            }
        }

        /// <summary>
        /// Mass Find and replace in presentation
        /// </summary>
        private void BatchTextReplace(Dictionary<string, string> replaces)
        {
            foreach (PowerPoint.Slide slide in this.presentation.Slides)
            {
                foreach(PowerPoint.Shape shape in slide.Shapes)
                {
                    if(shape.HasTextFrame == MsoTriState.msoTrue || shape.HasTable == MsoTriState.msoTrue)
                    {
                        //for textboxes
                        if (shape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            Util.ShapeTextReplace(shape.TextFrame.TextRange, replaces);
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
                                        Util.ShapeTextReplace(cell.Shape.TextFrame.TextRange, replaces);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Init slide Array
        /// </summary>
        /// <returns></returns>
        private object Slides(IDictionary<string, object> filter)
        {
            List<object> slides = new List<object>();

            foreach (PowerPoint.Slide pptSlide in this.presentation.Slides)
            {

                var slide = new Slide(pptSlide);
                if (slide.TestFilter(filter))
                {
                    slides.Add(slide.Invoke());
                }

            }

            return slides.ToArray();
        }

        private object PasteSlide(int index)
        {
            return new Slide(this.presentation.Slides.Paste(index).FirstOrDefault()).Invoke();
        }

        private object GetSelectedShape()
        {
            var selection = this.presentation.Application.ActiveWindow.Selection;
            if (selection.Type == PowerPoint.Enums.PpSelectionType.ppSelectionShapes)
            {
                return new Shape(selection.ShapeRange[1]).Invoke();
            }
            return null;
        }

        private object GetActiveSlide()
        {
            // Test if PowerPoint Presentation mode is currently running
            if (IsInSlideShowMode()){ 
                return new Slide(this.presentation.Application.SlideShowWindows[1].View.Slide).Invoke();
            }
            else {
                return new Slide(this.presentation.Application.ActiveWindow.Selection.SlideRange[1]).Invoke();
            }
        }

        /// <summary>
        /// Check if presentation in presentation mode
        /// </summary>
        private bool IsInSlideShowMode() {
            return (this.presentation.Application.SlideShowWindows.Count > 0);
        }

        /// <summary>
        /// 
        /// </summary>
        public PowerPoint.Presentation GetUnderlyingObject()
        {
            return this.presentation;
        }
        
        #region Properties

        public string Name
        {
            get
            {
                return this.presentation.Name;
            }
        }
        public string Path
        {
            get
            {
                return this.presentation.Path;
            }
        }
        public string FullName
        {
            get
            {
                return this.presentation.FullName;
            }
        }
        
        public float SlideHeight 
        {
            get
            {
                return this.presentation.PageSetup.SlideHeight;   
            }
        }
        
        public float SlideWidth 
        {
            get
            {
                return this.presentation.PageSetup.SlideWidth;   
            }
        }
        #endregion
    }
}
