using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
namespace ModelViewCallout_CSharp.csproj
{
    partial class SolidWorksMacro
    {
        ModelDoc2 swModelDoc;
        ModelDocExtension swModelDocExtn;
        ModelViewManager ViewMgr;
        ModelView swModelView;
        Callout Viewcallout;

        CalloutHandler handle = new CalloutHandler();

        public void RunMacro()
        {
            swModelDoc = (ModelDoc2)swApp.ActiveDoc;
            swModelDocExtn = swModelDoc.Extension;

            //ViewMgr = swModelDoc.ModelViewManager;
            //ViewMgr.ViewportDisplay = (int)swViewportDisplay_e.swViewportFourView;
            swModelDoc.GetModelViewCount();
            swModelView = (ModelView)swModelDoc.GetFirstModelView();
            while (((swModelView != null)))
            {
                Viewcallout = swModelView.CreateCallout(1, handle);
                Viewcallout.set_Label2(0, "TEST");
                Viewcallout.SkipColon = false;
                Viewcallout.set_ValueInactive(0, true);
                Viewcallout.SetTargetPoint(0, 0.0, 0.0, 0.0);
                Viewcallout.Display(true);
                System.Diagnostics.Debugger.Break();
                swModelView = (ModelView)swModelView.GetNext();
            }

        }

        public SldWorks swApp = (SldWorks)System.Runtime.InteropServices.Marshal.GetActiveObject("SldWorks.Application");

    }

    [ComVisibleAttribute(true)]
    public class CalloutHandler : SwCalloutHandler
    {

        public bool OnStringValueChanged(object pManipulator, int RowID, string Text)
        {

            Debug.Print("Text: " + Text);
            Debug.Print("Row: " + RowID);
            return true;
        }

    }
}