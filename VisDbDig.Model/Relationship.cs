using System.Collections.Generic;

namespace VisDbDig.Model
{ 
    public class Relationship
    {
        public string From { get; set; }
        public string To { get; set; }
        public bool OneToMany { get; set; }
    }
}