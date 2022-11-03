using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Entities
{
    public class Data
    {
        public string Component { get; set; }
        public string Parameter { get; set; }
        public List<string> Values { get; set; }
    }
    public class Element
    {
        public string Name { get; set; }
        public List<string> Values { get; set; }
        public List<Element> ElementChild { get; set; }
    }
}
