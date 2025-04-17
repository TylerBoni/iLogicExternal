using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Threading;
using Autodesk.iLogic.Interfaces;
using Inventor;
using Microsoft.Win32;
using System.Linq;

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

    /// <summary>
    /// Bridge class that implements the iLogic bridge functionality
    /// </summary>
    public class iLogicBridge
    {
        private bool isOpen = false;
        private dynamic iLogicAuto;
        private Inventor.Application inventorApp;
        private FileSystemWatcher watcher;
        private const string iLogicTransferFolder = "C:\\iLogicTransfer";
        private ConcurrentDictionary<string, bool> processingFiles = new ConcurrentDictionary<string, bool>();
        private bool isExportingRules = false;
        private readonly SynchronizationContext uiContext;
        private readonly object watcherLock = new object();
        private readonly ConcurrentDictionary<string, FileSystemWatcher> documentWatchers = new ConcurrentDictionary<string, FileSystemWatcher>();
        private bool isShuttingDown = false;

        public iLogicBridge(Inventor.Application app)
        {
            inventorApp = app;
            isOpen = true; // Since we're getting app from AddIn, it's already open
            uiContext = SynchronizationContext.Current ?? new SynchronizationContext();
        }

        public void Start()
        {
            try
            {
                // Check if the folder exists, and create it if it doesn't
                if (!System.IO.Directory.Exists(iLogicTransferFolder))
                {
                    System.IO.Directory.CreateDirectory(iLogicTransferFolder);
                }

                // Get iLogic
                iLogicAuto = GetiLogicAddIn(inventorApp);
                if (iLogicAuto != null)
                {
                    iLogicAuto.CallingFromOutside = false;

                    // Initial active document setup - do this on the UI thread
                    OpenDocumentRules();

                    // Set up document open event instead of activation event
                    inventorApp.ApplicationEvents.OnOpenDocument += ApplicationEvents_OnOpenDocument;
                    inventorApp.ApplicationEvents.OnCloseDocument += ApplicationEvents_OnCloseDocument;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in iLogic Bridge: {ex.Message}");
            }
        }

        public void Stop()
        {
            try
            {
                isShuttingDown = true;

                // Clean up event handlers
                if (inventorApp != null)
                {
                    inventorApp.ApplicationEvents.OnOpenDocument -= ApplicationEvents_OnOpenDocument;
                    inventorApp.ApplicationEvents.OnCloseDocument -= ApplicationEvents_OnCloseDocument;
                }

                // Clean up all watchers
                foreach (var watcherEntry in documentWatchers)
                {
                    try
                    {
                        var docWatcher = watcherEntry.Value;
                        if (docWatcher != null)
                        {
                            docWatcher.EnableRaisingEvents = false;
                            docWatcher.Changed -= OnChanged;
                            docWatcher.Created -= OnCreated;
                            docWatcher.Deleted -= OnDeleted;
                            docWatcher.Renamed -= OnRenamed;
                            docWatcher.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error disposing watcher: {ex.Message}");
                    }
                }
                documentWatchers.Clear();

                // Clean up iLogic transfer folder
                if (System.IO.Directory.Exists(iLogicTransferFolder))
                {
                    try
                    {
                        System.IO.Directory.Delete(iLogicTransferFolder, true);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error deleting transfer folder: {ex.Message}");
                    }
                }

                // Release references
                iLogicAuto = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error stopping iLogic Bridge: {ex.Message}");
            }
        }

        private void ApplicationEvents_OnOpenDocument(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;

            if (BeforeOrAfter == EventTimingEnum.kAfter)
            {
                // Log information about document opening
                System.Diagnostics.Debug.WriteLine($"Opening document: {System.IO.Path.GetFileName(FullDocumentName)}");

                // Use the UI synchronization context to handle document operations
                uiContext.Post(_ =>
                {
                    OpenDocumentRules(DocumentObject);
                }, null);
            }
        }

        private void ApplicationEvents_OnCloseDocument(_Document DocumentObject, string FullDocumentName, EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            HandlingCode = HandlingCodeEnum.kEventNotHandled;

            try
            {
                // We need to capture all the necessary information BEFORE the document is closed
                if (BeforeOrAfter == EventTimingEnum.kBefore)
                {
                    // Capture the information we need before the document is disposed
                    string docId = DocumentObject.InternalName.ToString();
                    string docFullName = FullDocumentName;

                    // Find the document folder by checking existing folders
                    string capturedFolder = null;
                    try
                    {
                        // Try to get the folder path using our existing methods
                        capturedFolder = GetDocumentFolder(DocumentObject);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error capturing document folder: {ex.Message}");
                    }

                    // Use the UI synchronization context to clean up after the document
                    uiContext.Post(_ =>
                    {
                        try
                        {
                            CleanupDocumentResources(docId, capturedFolder);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error in document cleanup: {ex.Message}");
                        }
                    }, null);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in OnCloseDocument: {ex.Message}");
            }
        }

        private void CleanupDocumentResources(string docId, string docFolder)
        {
            if (isShuttingDown) return;

            try
            {
                // Remove watcher for this document
                if (documentWatchers.TryRemove(docId, out FileSystemWatcher docWatcher))
                {
                    try
                    {
                        docWatcher.EnableRaisingEvents = false;
                        docWatcher.Changed -= OnChanged;
                        docWatcher.Created -= OnCreated;
                        docWatcher.Deleted -= OnDeleted;
                        docWatcher.Renamed -= OnRenamed;
                        docWatcher.Dispose();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error disposing document watcher: {ex.Message}");
                    }
                }

                // Clean up iLogic transfer folder - ensure it's deleted
                if (!string.IsNullOrEmpty(docFolder) && System.IO.Directory.Exists(docFolder))
                {
                    try
                    {
                        // Wait a moment to ensure all file operations are complete
                        System.Threading.Thread.Sleep(100);

                        // Attempt to delete the folder
                        System.IO.Directory.Delete(docFolder, true);
                        System.Diagnostics.Debug.WriteLine($"Successfully deleted folder for closed document: {docFolder}");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error deleting document folder: {ex.Message}");

                        // If the first attempt fails, try once more with each file
                        try
                        {
                            foreach (string file in System.IO.Directory.GetFiles(docFolder, "*.*", System.IO.SearchOption.AllDirectories))
                            {
                                try
                                {
                                    System.IO.File.Delete(file);
                                }
                                catch
                                {
                                    // Ignore errors on individual files
                                }
                            }

                            // Try to delete the folder again
                            System.IO.Directory.Delete(docFolder, true);
                            System.Diagnostics.Debug.WriteLine($"Successfully deleted folder on second attempt: {docFolder}");
                        }
                        catch (Exception retryEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"Failed to delete folder on retry: {retryEx.Message}");
                        }
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No folder found for document to clean up");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in CleanupDocumentResources: {ex.Message}");
            }
        }

        private void OpenDocumentRules(Inventor._Document documentObject = null)
        {
            if (isShuttingDown) return;

            // Create Inventor progress bar
            Inventor.ProgressBar progressBar = null;
            int totalSteps = 3;

            try
            {
                // Initialize progress bar
                if (inventorApp != null)
                {
                    progressBar = inventorApp.CreateProgressBar(false, totalSteps, "iLogic Rules");
                    progressBar.Message = "Processing rules...";
                    // First step
                    progressBar.UpdateProgress();
                }

                // Update the active document and export rules
                try
                {
                    if (inventorApp != null && iLogicAuto != null)
                    {
                        Inventor.Document doc = documentObject ?? inventorApp.ActiveDocument;
                        if (doc != null)
                        {
                            string docName = System.IO.Path.GetFileName(doc.FullDocumentName);
                            System.Diagnostics.Debug.WriteLine($"Setting up rules for {docName}...");

                            if (progressBar != null)
                            {
                                progressBar.Message = $"Creating file watcher for {docName}";
                                // Step is automatically incremented
                                progressBar.UpdateProgress();
                            }

                            // Create file watcher first
                            CreateFileWatcherForDocument(doc);

                            // Update progress
                            if (progressBar != null)
                            {
                                progressBar.Message = $"Exporting rules for {docName}";
                                // Step is automatically incremented  
                                progressBar.UpdateProgress();
                            }

                            System.Diagnostics.Debug.WriteLine($"Exporting rules for {docName}...");

                            // Then export rules
                            ExportRulesToFolder(doc);

                            // Final progress update
                            if (progressBar != null)
                            {
                                progressBar.Message = "Completed";
                                // Final step
                                progressBar.UpdateProgress();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error updating active document: {ex.Message}");
                    if (progressBar != null)
                    {
                        progressBar.Message = $"Error: {ex.Message}";
                    }
                }
            }
            finally
            {
                // Clean up progress bar
                if (progressBar != null)
                {
                    progressBar.Close();
                    progressBar = null;
                }

                // Log completion
                System.Diagnostics.Debug.WriteLine("Completed processing rules");
            }
        }

        private string GetDocumentFolder(Inventor.Document doc)
        {
            try
            {
                // Use document name instead of ID
                string docName = System.IO.Path.GetFileNameWithoutExtension(doc.FullDocumentName);
                // Sanitize the filename to ensure it's valid for folder names
                docName = string.Join("_", docName.Split(System.IO.Path.GetInvalidFileNameChars()));

                string docFolder = System.IO.Path.Combine(iLogicTransferFolder, docName);

                // If the folder already exists but isn't for this document, add a random suffix
                if (System.IO.Directory.Exists(docFolder))
                {
                    // Try to find a metadata file to see if it's the same document
                    string metaFile = System.IO.Path.Combine(docFolder, ".docid");
                    bool isSameDoc = false;

                    if (System.IO.File.Exists(metaFile))
                    {
                        try
                        {
                            string storedId = System.IO.File.ReadAllText(metaFile);
                            isSameDoc = storedId == doc.InternalName.ToString();
                        }
                        catch
                        {
                            // If we can't read the file, assume it's not the same document
                            isSameDoc = false;
                        }
                    }

                    if (!isSameDoc)
                    {
                        // Generate a random suffix of 4 characters
                        Random random = new Random();
                        string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
                        string suffix = "_" + new string(Enumerable.Repeat(chars, 4)
                            .Select(s => s[random.Next(s.Length)]).ToArray());

                        docFolder = System.IO.Path.Combine(iLogicTransferFolder, docName + suffix);

                        // In the unlikely event this folder also exists, keep trying
                        int attempts = 0;
                        while (System.IO.Directory.Exists(docFolder) && attempts < 10)
                        {
                            suffix = "_" + new string(Enumerable.Repeat(chars, 4)
                                .Select(s => s[random.Next(s.Length)]).ToArray());
                            docFolder = System.IO.Path.Combine(iLogicTransferFolder, docName + suffix);
                            attempts++;
                        }
                    }
                }

                // Create folder if it doesn't exist
                if (!System.IO.Directory.Exists(docFolder))
                {
                    System.IO.Directory.CreateDirectory(docFolder);

                    // Store document ID in metadata file for future reference
                    try
                    {
                        string metaFile = System.IO.Path.Combine(docFolder, ".docid");
                        System.IO.File.WriteAllText(metaFile, doc.InternalName.ToString());
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error creating metadata file: {ex.Message}");
                    }
                }

                return docFolder;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in GetDocumentFolder: {ex.Message}");
                return System.IO.Path.Combine(iLogicTransferFolder, "Temp_" + Guid.NewGuid().ToString().Substring(0, 8));
            }
        }

        private void ExportRulesToFolder(Inventor.Document doc)
        {
            // Prevent recursive calls or overlapping exports
            if (isExportingRules || isShuttingDown)
                return;

            try
            {
                isExportingRules = true;

                if (doc != null && iLogicAuto != null)
                {
                    // Get the document-specific folder
                    string docFolder = GetDocumentFolder(doc);

                    // Clear any existing rules in the folder
                    foreach (string file in System.IO.Directory.GetFiles(docFolder, "*.vb"))
                    {
                        try
                        {
                            System.IO.File.Delete(file);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error deleting file {file}: {ex.Message}");
                        }
                    }

                    // Export current document's rules
                    dynamic rules = iLogicAuto.get_Rules(doc);
                    if (rules == null)
                    {
                        System.Diagnostics.Debug.WriteLine("No rules found in the document.");
                        return;
                    }

                    foreach (iLogicRule r in rules)
                    {
                        try
                        {
                            System.IO.File.WriteAllText(System.IO.Path.Combine(docFolder, r.Name + ".vb"), r.Text);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error exporting rule {r.Name}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error exporting rules: {ex.Message}");
            }
            finally
            {
                isExportingRules = false;
            }
        }

        private Object GetiLogicAddIn(Inventor.Application app)
        {
            string iLogicGUID = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";
            Inventor.ApplicationAddIn iLogicAddIn = app.ApplicationAddIns.ItemById[iLogicGUID];
            iLogicAddIn.Activate();
            return iLogicAddIn.Automation;
        }

        private void CreateFileWatcherForDocument(Inventor.Document doc)
        {
            if (isShuttingDown) return;

            try
            {
                // Get document ID for watcher tracking
                string docId = doc.InternalName.ToString();

                // Get document-specific folder
                string docFolder = GetDocumentFolder(doc);

                // Don't create a new watcher if one exists for this document
                if (documentWatchers.ContainsKey(docId))
                {
                    // If the folder has changed, update the watcher path
                    FileSystemWatcher existingWatcher = documentWatchers[docId];
                    if (existingWatcher != null && existingWatcher.Path != docFolder)
                    {
                        try
                        {
                            existingWatcher.EnableRaisingEvents = false;
                            existingWatcher.Path = docFolder;
                            existingWatcher.EnableRaisingEvents = true;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error updating watcher path: {ex.Message}");
                            // If we can't update, remove and create a new one
                            if (documentWatchers.TryRemove(docId, out FileSystemWatcher oldWatcher))
                            {
                                try
                                {
                                    oldWatcher.EnableRaisingEvents = false;
                                    oldWatcher.Dispose();
                                }
                                catch { }
                            }
                        }
                    }
                    return;
                }

                // Create a new FileSystemWatcher and set its properties
                FileSystemWatcher newWatcher = new FileSystemWatcher
                {
                    Path = docFolder,
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.CreationTime,
                    Filter = "*.vb",
                    EnableRaisingEvents = false // Start disabled until we add handlers
                };

                // Add event handlers
                newWatcher.Changed += OnChanged;
                newWatcher.Created += OnCreated;
                newWatcher.Deleted += OnDeleted;
                newWatcher.Renamed += OnRenamed;

                // Store document reference with the watcher
                if (documentWatchers.TryAdd(docId, newWatcher))
                {
                    // Begin watching only after successfully adding to dictionary
                    newWatcher.EnableRaisingEvents = true;
                }
                else
                {
                    // If we couldn't add it to the dictionary, dispose it
                    newWatcher.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating file watcher: {ex.Message}");
            }
        }

        private void OnChanged(object source, FileSystemEventArgs e)
        {
            if (isShuttingDown) return;
            ProcessFileChange(e.FullPath, FileChangeType.Changed);
        }

        private void OnCreated(object source, FileSystemEventArgs e)
        {
            if (isShuttingDown) return;
            ProcessFileChange(e.FullPath, FileChangeType.Created);
        }

        private void OnDeleted(object source, FileSystemEventArgs e)
        {
            if (isShuttingDown) return;
            ProcessFileChange(e.FullPath, FileChangeType.Deleted);
        }

        private void OnRenamed(object source, RenamedEventArgs e)
        {
            if (isShuttingDown) return;

            // Skip temp files
            if (e.OldName + "~" == e.Name)
            {
                System.Diagnostics.Debug.WriteLine("**VIM SWAP FILE, NOT A RENAME**");
                return;
            }

            ProcessFileChange(e.FullPath, FileChangeType.Renamed, e.OldFullPath);
        }

        private enum FileChangeType
        {
            Changed,
            Created,
            Deleted,
            Renamed
        }

        private void ProcessFileChange(string fullPath, FileChangeType changeType, string oldFullPath = null)
        {
            if (isShuttingDown) return;

            // Prevent processing the same file multiple times concurrently
            string key = fullPath + changeType.ToString();
            if (!processingFiles.TryAdd(key, true))
                return;

            try
            {
                // Use the UI sync context to ensure we're on the right thread for Inventor API calls
                uiContext.Post(_ =>
                {
                    try
                    {
                        if (isShuttingDown || inventorApp == null || iLogicAuto == null)
                        {
                            bool temp;
                            processingFiles.TryRemove(key, out temp);
                            return;
                        }

                        Inventor.Document doc = null;
                        try
                        {
                            doc = inventorApp.ActiveDocument;
                            if (doc == null)
                            {
                                bool temp;
                                processingFiles.TryRemove(key, out temp);
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error getting active document: {ex.Message}");
                            bool temp;
                            processingFiles.TryRemove(key, out temp);
                            return;
                        }

                        string fileName = System.IO.Path.GetFileName(fullPath);
                        string ruleName = System.IO.Path.GetFileNameWithoutExtension(fileName);

                        try
                        {
                            switch (changeType)
                            {
                                case FileChangeType.Changed:
                                    try
                                    {
                                        dynamic rule = iLogicAuto.GetRule(doc, ruleName);
                                        if (rule != null)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"Updating Rule {ruleName}...");
                                            string newText = System.IO.File.ReadAllText(fullPath);
                                            rule.Text = newText;
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"**FAILED TO WRITE TO {ruleName} ON CHANGE: RULE DOESNT EXIST**");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Error updating rule {ruleName}: {ex.Message}");
                                    }
                                    break;

                                case FileChangeType.Created:
                                    try
                                    {
                                        dynamic existingRule = iLogicAuto.GetRule(doc, ruleName);
                                        if (existingRule == null)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"Creating Rule {ruleName}...");
                                            string newText = System.IO.File.ReadAllText(fullPath);
                                            dynamic newRule = iLogicAuto.AddRule(doc, ruleName, "");
                                            newRule.Text = newText;
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"**RULE ALREADY EXISTS: {ruleName}**");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Error creating rule {ruleName}: {ex.Message}");
                                    }
                                    break;

                                case FileChangeType.Deleted:
                                    try
                                    {
                                        dynamic ruleToDelete = iLogicAuto.GetRule(doc, ruleName);
                                        if (ruleToDelete != null)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"WARNING: Removing Rule {ruleName}...");
                                            iLogicAuto.DeleteRule(doc, ruleName);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Error deleting rule {ruleName}: {ex.Message}");
                                    }
                                    break;

                                case FileChangeType.Renamed:
                                    try
                                    {
                                        string oldFileName = System.IO.Path.GetFileName(oldFullPath);
                                        string oldRuleName = System.IO.Path.GetFileNameWithoutExtension(oldFileName);

                                        dynamic oldRule = iLogicAuto.GetRule(doc, oldRuleName);
                                        if (oldRule != null)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"Renaming Rule {oldRuleName} To {ruleName}...");
                                            string ruleText = oldRule.Text;
                                            iLogicAuto.DeleteRule(doc, oldRuleName);
                                            dynamic renamedRule = iLogicAuto.AddRule(doc, ruleName, "");
                                            renamedRule.Text = ruleText;
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debug.WriteLine($"**OLD RULE {oldRuleName} DOES NOT EXIST**");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Error renaming rule: {ex.Message}");
                                    }
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error processing {changeType} for {ruleName}: {ex.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error processing file change: {ex.Message}");
                    }
                    finally
                    {
                        // Remove from processing dictionary when done
                        bool temp;
                        processingFiles.TryRemove(key, out temp);
                    }
                }, null);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error initiating file change processing: {ex.Message}");
                bool temp;
                processingFiles.TryRemove(key, out temp);
            }
        }
    }
}


