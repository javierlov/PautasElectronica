using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PautasPublicidad.DTO
{
    [AttributeUsage(AttributeTargets.Class |
    AttributeTargets.Constructor |
    AttributeTargets.Field |
    AttributeTargets.Method |
    AttributeTargets.Property,
    AllowMultiple = true)]
    public class SortFieldAttribute : System.Attribute
    {
        public string Property { get; set; }
        //public string Property
        //{
        //    get
        //    {
        //        return Property;
        //    }
        //    set
        //    {
        //        Property = value;
        //    }
        //}
    }
}
