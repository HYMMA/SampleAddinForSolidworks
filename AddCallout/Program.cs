using ModelViewCallout_CSharp.csproj;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddCallout
{
    public class Program
    {
        public static void Main()
        {
            var macro = new SolidWorksMacro();
            macro.RunMacro();
        }
    }
}
