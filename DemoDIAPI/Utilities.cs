using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DemoDIAPI
{
    public static class Utilities
    {
        /// <summary>
        /// Releases the COM objects.
        /// </summary>
        /// <param name="objects">COM objects to release.</param>
        public static void Release(params object[] objects)
        {
            foreach (var obj in objects)
            {
                ReleaseOne(obj);
            }
        }

        /// <summary>
        /// Checks if the object is a COM object.
        /// </summary>
        /// <param name="o">The object to check.</param>
        /// <returns>True if the object is a COM object, false otherwise.</returns>
        private static bool NotComObj(object o)
        {
            return !"System.__ComObject".Equals(o.GetType().ToString());
        }

        /// <summary>
        /// Releases the COM object.
        /// </summary>
        /// <param name="o">The object to release.</param>
        private static void ReleaseOne(object o)
        {
            if (o == null || NotComObj(o))
            {
                return;
            }

            Marshal.ReleaseComObject(o);
        }
    }
}
