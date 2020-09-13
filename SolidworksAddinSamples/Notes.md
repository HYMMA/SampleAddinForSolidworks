# Working with event handlers in Solidworks Addins
Engineers who use better tools tend to achieve better results. As a mechanical engineer I have always been fascinated by automation in design and how it can save hours of tedious work. So I decided to create a solidworks add-in to use at my work. However I realized things could become complex very fast. Although there are sample codes such as solidworks SDK to use as a template, there is not good documentation for it. In this article I will address some of the pitfalls in *Event handling* for a solidworks add-in. 
Events in programming world are such a powerful tool in the hands of the informed designer. However sample codes in solidworks SDK are not easy to understand and reason about. I will use a sample project here that you can download **from this link on GitHub.** To read through this article you need to be comfortable with the concept of event in .NET environment. If you decided to learn more about it here is a **good article about events in .NET**

## First things first
Before we continue I'd like to point out to a few important things. In solidworks some interfaces do not expose any events. But the interfaces that query-interface to them (Derive from them) do. For example *IModelDoc* is the parent of *ModelDoc* but only the child exposes the events! So we should use casting and define a helper object ro access the events as shown below.

```C#
ModelDoc2 document = (ModelDoc2)addin.GetFirstDocument();
IModelDoc myModelAbstraction = document.GetFirstModelView();
ModelDoc myModel= (ModelDoc)myModelAbstraction;
myModel.eventName = //Attach your event handlers via delegates
```

Secondly, we should subscribe to events in three levels of our addin
1. Addin level
2. Document level
3. Model-view level

In all of these steps we use a key-value pair to register the objects whose events are being listened to. 

>The class hierarchy in solidworks is: ISldWorks -> IModelDoc -> IModelView
>to access IModelView objects you need to access *IModelDoc* first and to access *IModelDoc* you should use *ISldWorks* which is the addin itself.

>we use two key-value pair objects (hashtables) where the key is the object you want to subscribe to and the value is the helper class that makes subscription happen. We define two helper classes. One is used for documents and the other is used for Model-views.

## Attaching Event Handlers to Addin 
Once the application is loaded we call **AttachSwEvents** method from *ConnectToSw* so that we can listen to events like *FileOpenPostNotify*. This way when user opens a file we (our addin) get notified and call our even handlers to do something. But there are two different scenarios that we should be aware of.
1. when user opens some documents in solidworks and then starts our addin
2. when user starts our addin first and then opens some documents in solidworks

>**AttachSwEvents** will attach event handlers to the addin. We know that addin object is derived from *ISldWorks*. But *ISldWorks* interface does not expose any events so we should use casting to access the events in solidworks. `addinWithEvents = (SolidWorks.Interop.sldworks.SldWorks)addin;`

When user runs the addin and opens a file then the *FileOpenPostNotify* event will trigger the event handlers. But in case he decided to open some documents first and then run the addin we should iterate through all open documents first. This way we have a mechanism in the software to listen to all the events in solidworks documents. The **AttachEventsToAllDocuments** method will iterate through open documents and subscribes our event handlers to the documents and to their *ModleViews*.

>ModelView class is the view for each document in solidworks. It is possible to get more than one view of the same document (part, assembly).

## Attaching event handlers to documents and Model-views
To access events that happen in the document itself we should use a *ModelDoc2* object. We access this object from the addin instance.

```C#
 ModelDoc2 ThisDocument = (ModelDoc2)addin.GetFirstDocument();
 ```

As mentioned earlier **AttachEventsToAllDocuments** will iterate through documents calls *AttachEventHandlersToDocument* on each one of them. *AttachEventHandlersToDocument* uses a *Hashtable* key-value pair to register the documents that have been processed. In this hashtable the keys are the document objects and the values are instances of **DocumentEventHandler** class.

### DocumentEventHandler.cs
This is a helper class that handles events for the documents and their model-views. There are three different classes that derive from this class. Each one handles events of a specific type of document.

```C#
//handles events for Part documents
public class PartEventHandler : DocumentEventHandler{} 

 //handles events for Assembly documents
public class AssemblyEventHandler : DocumentEventHandler{}
 
 //handles events for Drawing documents
public class DrawingEventHandler : DocumentEventHandler{}
```
In *DocumentEventHandler* you can find these methods and properties.

**ConnectModelViews()**
This method uses a hashtable field called *openModelViews* and iterates through modelViews. In this hashtable the keys are the *modelView* and the values are *ModelViewHelper*.


**DisconnectModelViews**
This method closes events on all currently open documents. Firstly copies the modelViews from the hashtable field to an array of objects. Then iterate through the array and detaches the event handlers from each one. 

```C#
public bool DetachEventHandlers()
        {
            modelView.DestroyNotify2 -= new DModelViewEvents_DestroyNotify2EventHandler(OnDestroy);
            modelView.RepaintNotify -= new DModelViewEvents_RepaintNotifyEventHandler(OnRepaint);
            parent.DetachModelViewEventHandler(modelView);
            return true;
        }
```
         
### ModelViewHelper.cs
This is a helper class to attach event handlers to *ModelView* objects. Let us take a look at the constructor of the class

```C#
public ModelViewHelper(SwAddin addin, IModelView mv, DocumentEventHandler doc)
		{
			userAddin = addin;
			modelView = (ModelView)mv;
			iSwApp = (ISldWorks)userAddin.SwApp;
			parent = doc;
		}
```
It needs a *SwAddin* and *IModelView* and *DocumentEventHandler* object. This class is called from *DocumentEventHandler.ConnectModelViews()* and exposes the events of *ModelView* objects. Remember there was a hashtable to register documents with their events? In that hashtable the keys are the *modelView* and the values are *ModelViewHelper*. 

## Summary
In summary, different solidworks components have their own events that we access from their objects. For example to respond to a file open event we use *FileOpenPostNotify* event from a *SldWorks* object. Similarly to trigger an action when a document is closed we use *DestroyNotify2* from document's *ModelView* object. Later when the user closes the software or the documents these events need to be unloaded. So we use two key-value pair objects (hashtables) where the key is the object you want to subscribe to and the value is the helper class that makes subscription happen. We define two helper classes. One is used for documents and the other is used for Model-views. The later is accessed from within the document helper class which in turn is accessed from main add-in object. In other words, once the addin starts we access the documents events via a helper class which itself uses another helper class to access the events of each document's model-views. When user closes a mode-view we want to un-subscribe from its events so we use `modelView.DestroyNotify2 -= new DModelViewEvents_DestroyNotify2EventHandler(OnDestroy);` and remove that object from its hashtable.
