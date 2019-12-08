
using System.Collections.Generic;
using System.Text;

namespace VisDbDig.Model
{
    public class Field
    {
        public string Name { get; set; }
        public string DataType { get; set; }
        public bool OneToMany { get; set; }
    }
}