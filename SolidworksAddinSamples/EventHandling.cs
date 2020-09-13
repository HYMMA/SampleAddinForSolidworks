using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections;

namespace HymmaSampleAddin
{
    public class DocumentEventHandler
    {
        protected ISldWorks addin;
        protected ModelDoc2 document;
        protected SwAddin userAddin;

        protected Hashtable openModelViews;

        public DocumentEventHandler(ModelDoc2 document, SwAddin addin)
        {
            this.document = document;
            userAddin = addin;
            this.addin = (ISldWorks)userAddin.SwApp;
            openModelViews = new Hashtable();
        }

        virtual public bool AttachEventHandlers()
        {
            return true;
        }

        virtual public bool DetachEventHandlers()
        {
            return true;
        }
        /// <summary>
        /// uses a hashtable field called openModelViews and iterates through modelViews of this document
        ///if the hashtable openModelViews does not have a modelView
        ///creates a DocumentView object and adds EventHandlers (OnDestroyNotify and OnRepaintNotify) to it and then
        ///adds the modelView and documentView to the hashtable
        /// </summary>
        /// <returns>true</returns>
        public bool ConnectModelViews()
        {
            IModelView modelView;
            modelView = (IModelView)document.GetFirstModelView();
            while (modelView != null)
            {
                if (!openModelViews.Contains(modelView))
                {
                    ModelViewHelper modelViewHelper = new ModelViewHelper(userAddin, modelView, this);
                    modelViewHelper.AttachEventHandlers();
                    openModelViews.Add(modelView, modelViewHelper);
                }
                modelView = (IModelView)modelView.GetNext();
            }
            return true;
        }

        public bool DisconnectModelViews()
        {
            //Close events on all currently open docs
            ModelViewHelper modelViewHelper;
            int kyeQuantity;
            kyeQuantity = openModelViews.Count;
            if (kyeQuantity == 0)
            {
                return false;
            }
            object[] modelViews = new object[kyeQuantity];

            //Remove all ModelView event handlers
            openModelViews.Keys.CopyTo(modelViews, 0);
            foreach (ModelView modelView in modelViews)
            {
                modelViewHelper = (ModelViewHelper)openModelViews[modelView];
                modelViewHelper.DetachEventHandlers();
                openModelViews.Remove(modelView);
                modelViewHelper = null;
            }
            return true;
        }

        /// <summary>
        /// we cannot use this code inside of DisconnectModelViews because we use modelView to iterate through the hastable
        /// </summary>
        /// <param name="modelView">represents the model view for a solidworks document</param>
        /// <returns></returns>
        public bool DetachModelViewEventHandler(ModelView modelView)
        {
            ModelViewHelper modelViewHelper;
            if (openModelViews.Contains(modelView))
            {
                modelViewHelper = (ModelViewHelper)openModelViews[modelView];
                openModelViews.Remove(modelView);
                modelView = null;
                modelViewHelper = null;
            }
            return true;
        }
    }

    public class PartEventHandler : DocumentEventHandler
    {
        PartDoc doc;

        public PartEventHandler(ModelDoc2 modDoc, SwAddin addin)
            : base(modDoc, addin)
        {
            doc = (PartDoc)document;
        }

        override public bool AttachEventHandlers()
        {
            doc.DestroyNotify += new DPartDocEvents_DestroyNotifyEventHandler(OnDestroy);
            doc.NewSelectionNotify += new DPartDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);
            ConnectModelViews();
            return true;
        }

        override public bool DetachEventHandlers()
        {
            doc.DestroyNotify -= new DPartDocEvents_DestroyNotifyEventHandler(OnDestroy);
            doc.NewSelectionNotify -= new DPartDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);
            DisconnectModelViews();
            userAddin.RemoveModelFromDocEventRepo(document);
            return true;
        }
        //Event Handlers
        public int OnDestroy()
        {
            DetachEventHandlers();
            return 0;
        }

        public int OnNewSelection()
        {
            return 0;
        }
    }

    public class AssemblyEventHandler : DocumentEventHandler
    {
        AssemblyDoc assemblyDocument;
        SwAddin swAddin;

        public AssemblyEventHandler(ModelDoc2 modDoc, SwAddin addin)
            : base(modDoc, addin)
        {
            assemblyDocument = (AssemblyDoc)document;
            swAddin = addin;
        }

        override public bool AttachEventHandlers()
        {
            assemblyDocument.DestroyNotify += new DAssemblyDocEvents_DestroyNotifyEventHandler(OnDestroy);
            assemblyDocument.NewSelectionNotify += new DAssemblyDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);
            assemblyDocument.ComponentStateChangeNotify2 += new DAssemblyDocEvents_ComponentStateChangeNotify2EventHandler(ComponentStateChangeNotify2);
            assemblyDocument.ComponentStateChangeNotify += new DAssemblyDocEvents_ComponentStateChangeNotifyEventHandler(ComponentStateChangeNotify);
            assemblyDocument.ComponentVisualPropertiesChangeNotify += new DAssemblyDocEvents_ComponentVisualPropertiesChangeNotifyEventHandler(ComponentVisualPropertiesChangeNotify);
            assemblyDocument.ComponentDisplayStateChangeNotify += new DAssemblyDocEvents_ComponentDisplayStateChangeNotifyEventHandler(ComponentDisplayStateChangeNotify);
            ConnectModelViews();
            return true;
        }

        override public bool DetachEventHandlers()
        {
            assemblyDocument.DestroyNotify -= new DAssemblyDocEvents_DestroyNotifyEventHandler(OnDestroy);
            assemblyDocument.NewSelectionNotify -= new DAssemblyDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);
            assemblyDocument.ComponentStateChangeNotify2 -= new DAssemblyDocEvents_ComponentStateChangeNotify2EventHandler(ComponentStateChangeNotify2);
            assemblyDocument.ComponentStateChangeNotify -= new DAssemblyDocEvents_ComponentStateChangeNotifyEventHandler(ComponentStateChangeNotify);
            assemblyDocument.ComponentVisualPropertiesChangeNotify -= new DAssemblyDocEvents_ComponentVisualPropertiesChangeNotifyEventHandler(ComponentVisualPropertiesChangeNotify);
            assemblyDocument.ComponentDisplayStateChangeNotify -= new DAssemblyDocEvents_ComponentDisplayStateChangeNotifyEventHandler(ComponentDisplayStateChangeNotify);
            DisconnectModelViews();

            userAddin.RemoveModelFromDocEventRepo(document);
            return true;
        }

        //Event Handlers
        public int OnDestroy()
        {
            DetachEventHandlers();
            return 0;
        }

        public int OnNewSelection()
        {
            return 0;
        }

        //attach events to a component if it becomes resolved
        protected int ComponentStateChange(object componentModel, short newCompState)
        {
            ModelDoc2 modDoc = (ModelDoc2)componentModel;
            swComponentSuppressionState_e newState = (swComponentSuppressionState_e)newCompState;


            switch (newState)
            {

                case swComponentSuppressionState_e.swComponentFullyResolved:
                    {
                        if ((modDoc != null) & !this.swAddin.OpenDocs.Contains(modDoc))
                        {
                            this.swAddin.AttachEventHandlersToDocument(modDoc);
                        }
                        break;
                    }

                case swComponentSuppressionState_e.swComponentResolved:
                    {
                        if ((modDoc != null) & !this.swAddin.OpenDocs.Contains(modDoc))
                        {
                            this.swAddin.AttachEventHandlersToDocument(modDoc);
                        }
                        break;
                    }

            }
            return 0;
        }

        protected int ComponentStateChange(object componentModel)
        {
            ComponentStateChange(componentModel, (short)swComponentSuppressionState_e.swComponentResolved);
            return 0;
        }


        public int ComponentStateChangeNotify2(object componentModel, string CompName, short oldCompState, short newCompState)
        {
            return ComponentStateChange(componentModel, newCompState);
        }

        int ComponentStateChangeNotify(object componentModel, short oldCompState, short newCompState)
        {
            return ComponentStateChange(componentModel, newCompState);
        }

        int ComponentDisplayStateChangeNotify(object swObject)
        {
            Component2 component = (Component2)swObject;
            ModelDoc2 modDoc = (ModelDoc2)component.GetModelDoc();

            return ComponentStateChange(modDoc);
        }

        int ComponentVisualPropertiesChangeNotify(object swObject)
        {
            Component2 component = (Component2)swObject;
            ModelDoc2 modDoc = (ModelDoc2)component.GetModelDoc();

            return ComponentStateChange(modDoc);
        }




    }

    public class DrawingEventHandler : DocumentEventHandler
    {
        DrawingDoc doc;

        public DrawingEventHandler(ModelDoc2 modDoc, SwAddin addin)
            : base(modDoc, addin)
        {
            doc = (DrawingDoc)document;
        }

        override public bool AttachEventHandlers()
        {
            doc.DestroyNotify += new DDrawingDocEvents_DestroyNotifyEventHandler(OnDestroy);
            doc.NewSelectionNotify += new DDrawingDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);

            ConnectModelViews();

            return true;
        }

        override public bool DetachEventHandlers()
        {
            doc.DestroyNotify -= new DDrawingDocEvents_DestroyNotifyEventHandler(OnDestroy);
            doc.NewSelectionNotify -= new DDrawingDocEvents_NewSelectionNotifyEventHandler(OnNewSelection);

            DisconnectModelViews();

            userAddin.RemoveModelFromDocEventRepo(document);
            return true;
        }

        //Event Handlers
        public int OnDestroy()
        {
            DetachEventHandlers();
            return 0;
        }

        public int OnNewSelection()
        {
            return 0;
        }
    }

    /// <summary>
    /// It is a helper class that implements MovdelView instead of IMovdelView so that it can access the events in that object
    /// this is achieved via casting of the IMovdelView object into a MovdelView
    /// </summary>
    public class ModelViewHelper
    {
        ISldWorks iSwApp;
        SwAddin userAddin;
        ModelView modelView;
        DocumentEventHandler parent;

        public ModelViewHelper(SwAddin addin, IModelView mv, DocumentEventHandler doc)
        {
            userAddin = addin;
            modelView = (ModelView)mv;
            iSwApp = (ISldWorks)userAddin.SwApp;
            parent = doc;
        }

        public bool AttachEventHandlers()
        {
            modelView.DestroyNotify2 += new DModelViewEvents_DestroyNotify2EventHandler(OnDestroy);
            modelView.RepaintNotify += new DModelViewEvents_RepaintNotifyEventHandler(OnRepaint);
            return true;
        }

        public bool DetachEventHandlers()
        {
            modelView.DestroyNotify2 -= new DModelViewEvents_DestroyNotify2EventHandler(OnDestroy);
            modelView.RepaintNotify -= new DModelViewEvents_RepaintNotifyEventHandler(OnRepaint);
            parent.DetachModelViewEventHandler(modelView);
            return true;
        }

        //EventHandlers
        public int OnDestroy(int destroyType)
        {
            switch (destroyType)
            {
                case (int)swDestroyNotifyType_e.swDestroyNotifyHidden:
                    return 0;

                case (int)swDestroyNotifyType_e.swDestroyNotifyDestroy:
                    return 0;
            }

            return 0;
        }

        public int OnRepaint(int repaintType)
        {
            return 0;
        }
    }

}
