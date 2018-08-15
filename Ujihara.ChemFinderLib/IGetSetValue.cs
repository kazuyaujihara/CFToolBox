using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ujihara.Chemistry
{
    public interface IGetSetValue
    {
        object GetValue(string key);
        void SetValue(string key, object value);
    }
}
