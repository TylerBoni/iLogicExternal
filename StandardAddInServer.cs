using System;
using System.Runtime.InteropServices;
using Inventor;
using iLogicExternal;

namespace iLogicExternal
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("6cfad819-5a80-4533-b520-a8356472404b")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {
        // Inventor application object.
        private Inventor.Application app;
        private iLogicBridge bridge;

        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            // Initialize AddIn members.
            app = addInSiteObject.Application;

            // Initialize and start the iLogic Bridge
            bridge = new iLogicBridge(app);
            bridge.Start();
        }

        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // Stop the iLogic Bridge
            if (bridge != null)
            {
                bridge.Stop();
                bridge = null;
            }

            // Release objects.
            app = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                // Return the bridge object as automation
                return bridge;
            }
        }

        #endregion
    }
}


