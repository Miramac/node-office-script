using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeScript.Report;

namespace OfficeScript
{
    class Startup
    {
        public async Task<object> PowerPointApplication(dynamic options)
        {
            return new PowerPointApplication().Invoke();
        }

    }
}
