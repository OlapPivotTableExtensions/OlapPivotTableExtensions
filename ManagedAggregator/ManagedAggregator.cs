using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;

namespace ManagedHelpers
{
    // This interface will be implemented by the outer object in the
    // aggregation - that is, by the shim.
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("7B70C487-B741-4973-B915-B812A91BDF63")]
    internal interface IComAggregator
    {
        void SetInnerPointer(IntPtr pUnkInner);
    }

    // This interface is implemented by the managed aggregator - the single
    // method is a wrapper around Marshal.CreateAggregatedObject, which can be
    // called from unmanaged code (that is, called from the shim).
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("142A261B-1550-4849-B109-715AA4629A14")]
    internal interface IManagedAggregator
    {
        void CreateAggregatedInstance(
            string assemblyName, string typeName, IComAggregator outerObject);
    }

    // The unmanaged shim will instantiate this object in order to call
    // through to Marshal.CreateAggregatedObject.
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("ManagedHelpers.ManagedAggregator")]
    internal class ManagedAggregator : IManagedAggregator
    {
        public void CreateAggregatedInstance(
            string assemblyName, string typeName, IComAggregator outerObject)
        {
            IntPtr pOuter = IntPtr.Zero;
            IntPtr pInner = IntPtr.Zero;

            try
            {
                // We use Marshal.CreateAggregatedObject to create a CCW where
                // the inner object (the target managed add-in) is aggregated 
                // with the supplied outer object (the shim).
                pOuter = Marshal.GetIUnknownForObject(outerObject);
                object innerObject =
                    AppDomain.CurrentDomain.CreateInstanceAndUnwrap(
                    assemblyName, typeName);
                pInner = Marshal.CreateAggregatedObject(pOuter, innerObject);

                // Make sure the shim has a pointer to the add-in.
                outerObject.SetInnerPointer(pInner);
            }
            finally
            {
                if (pOuter != IntPtr.Zero)
                {
                    Marshal.Release(pOuter);
                }
                if (pInner != IntPtr.Zero)
                {
                    Marshal.Release(pInner);
                }

                // FIX: Bug discovered after release of 2.3.1.0.
                // We call ReleaseComObject on the outer object (ConnectProxy)
                // to make sure we delete the RCW, and prevent the CLR from
                // holding onto it indefinitely (and keeping the host alive).
                Marshal.ReleaseComObject(outerObject);
            }
        }
    }

}
