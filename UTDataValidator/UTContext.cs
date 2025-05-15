using System;
using System.Collections.Generic;

namespace UTDataValidator
{
    public class UTContext<T>
    {
        public T OutputValue { get; set; }
        public List<string> ErrorMessages { get; set; }
    }
}