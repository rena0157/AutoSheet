using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;

namespace ACadLib.Utilities
{
    public class ACadLogger
    {
        private readonly Document _activeDocument;

        public ACadLogger(Document document)
        {
            _activeDocument = document;
        }

        public void Log(object message)
        {
            _activeDocument.Editor.WriteMessage($"\n[{DateTime.Now}] - {message}");
        }
    }
}
