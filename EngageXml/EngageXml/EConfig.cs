using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace EngageXml
{
    internal class EConfig
    {
        XDocument config;
        public EConfig(string path) {
            config = XDocument.Load(path);
        }
        public IEnumerable<XElement> ParamPatches { get {
                return config.Root.Element("ParamPatches").Elements("Patch");
            } 
        }

        public IEnumerable<XElement> FilePatches
        {
            get
            {
                return config.Root.Element("FilePatches").Elements("File");
            }
        }
    }
}
