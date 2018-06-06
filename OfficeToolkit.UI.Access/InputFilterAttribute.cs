using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeToolkit.Common
{
    [AttributeUsage( AttributeTargets.Class, AllowMultiple=true, Inherited=true)]
    public class InputFilterAttribute : Attribute
    {
        public InputFilterAttribute(string extension, string description)
        {
            this.Extension = extension;
            this.Description = description;
        }

        public readonly string Extension;
        public readonly string Description;
    }
}
