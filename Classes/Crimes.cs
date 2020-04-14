using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab2.Classes
{
    public class Crimes
    {
        private string id;
        private string name;
        private string discription;
        private string sourse;
        private string impactedobject;
        private string confidentiality;
        private string integrity;
        private string accessibility;
        public string Id
        {
            get { return id; }
            set
            {
                if (Int32.Parse(value) > 0 & Int32.Parse(value) < 10)
                {
                    id = "УБИ.00" + value;
                }
                else if (Int32.Parse(value) > 9 & Int32.Parse(value) < 100)
                {
                    id = "УБИ.0" + value;
                }
                else
                {
                    id = "УБИ." + value;
                }
            }
        }
        public string Name { get; set; }
        public string Discription { get; set; }
        public string Source { get; set; }
        public string ImpactedObject{ get; set; }
        public string Confidentiality { get { return confidentiality; } set { if (value == "1") confidentiality = "Да"; else confidentiality = "Нет"; } } // yes/no
        public string Integrity { get { return integrity; } set { if (value == "1") integrity = "Да"; else integrity = "Нет"; } }// yes/no
        public string Accessibility { get { return accessibility; } set { if (value == "1") accessibility = "Да"; else accessibility  = "Нет"; } } // yes/no

        public Crimes(string id, string name, string discription, string source, string impactedObject, string confidentiality, string integrity, string accessibility)
        {
            Id = id;
            Name = name;
            Discription = discription;
            Source = source;
            ImpactedObject = impactedObject;
            Confidentiality = confidentiality;
            Integrity = integrity;
            Accessibility = accessibility;
        }
    }
}
