using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using System.Drawing.Design;
using System.ComponentModel.Design;
using System.Reflection;
using System.Resources;
using System.Runtime.Remoting;

namespace Firefly.Box.VSIntegration
{
    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    ///
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane, 
    /// usually implemented by the package implementer.
    ///
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its 
    /// implementation of the IVsWindowPane interface.
    /// </summary>
    [Guid("8f28e43a-e6c8-471b-8d1b-7af0ce85884f")]
    public class FireflyComponentBrowser : ToolWindowPane
    {
        // This is the user control hosted by the tool window; it is exposed to the base class 
        // using the Window property. Note that, even if this class implements IDispose, we are
        // not calling Dispose on this object. This is because ToolWindowPane calls Dispose on 
        // the object returned by the Window property.
        private MyControl control;

        System.Type _lastDesignedType = typeof(object);

        public FireflyComponentBrowser()
            :
            base(null)
        {
            // Set the window title reading it from the resources.
            this.Caption = "Firefly Toolbox";
            // Set the image that will appear on the tab of the window frame
            // when docked with an other window
            // The resource ID correspond to the one defined in the resx file
            // while the Index is the offset in the bitmap strip. Each image in
            // the strip being 16x16.
            this.BitmapResourceID = 301;
            this.BitmapIndex = 1;

            control = new MyControl();

            Panel p = new Panel();
            p.BorderStyle = BorderStyle.None;
            p.Dock = DockStyle.Fill;
            p.BackColor = SystemColors.ControlDark;

            System.Windows.Forms.TreeView treeView = new TreeView();
            ImageListManager imageList = new ImageListManager();
            imageList.SetToTreeView(treeView);


            //treeView.Dock = DockStyle.Fill;
            treeView.Anchor = AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
            treeView.BorderStyle = BorderStyle.None;
            uint cookie = 0;
            Cmd addSelectedControlToToolBox =
                delegate
                {
                    try
                    {
                        IToolboxService toolboxService = (IToolboxService)this.GetService(typeof(IToolboxService));
                        toolboxService.AddToolboxItem((ToolboxItem)treeView.SelectedNode.Tag,
                                                      toolboxService.SelectedCategory ?? "General");
                    }
                    catch
                    {
                    }
                };
            ToolStrip toolStrip = new ToolStrip();
            toolStrip.GripStyle = ToolStripGripStyle.Hidden;
            //toolStrip.Dock = DockStyle.Top;
            ProfessionalColorTable professionalColorTable = new ProfessionalColorTable();
            professionalColorTable.UseSystemColors = true;
            ToolStripProfessionalRenderer toolStripRenderer = new ToolStripProfessionalRenderer(professionalColorTable);
            toolStripRenderer.RoundedEdges = false;
            toolStrip.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            toolStrip.Renderer = toolStripRenderer;

            control.Controls.Add(p);
            toolStrip.AutoSize = false;
            toolStrip.Bounds = new Rectangle(1, 0, p.Width - 2, toolStrip.Height + 1);
            treeView.Bounds = new Rectangle(1, toolStrip.Height + 1, p.Width - 2, p.Height - toolStrip.Height - 2);
            //            p.Padding = new Padding(1,0,1,1);
            //            treeView.Margin = new Padding(0, 1, 0, 0);
            p.Controls.Add(treeView);
            p.Controls.Add(toolStrip);

            imageList.Add("Assembly", Resources.Assembly);
            imageList.Add("Namespace", Resources.Namespace);

            Bitmap bmp = Resources.AddToFavoritesHS;
            ToolStripButton addToToolBoxButton = new ToolStripButton(bmp);
            addToToolBoxButton.ToolTipText = "Add to tool box";
            addToToolBoxButton.ImageTransparentColor = System.Drawing.Color.Black;
            toolStrip.Items.Add(addToToolBoxButton);
            addToToolBoxButton.Click +=
                delegate { addSelectedControlToToolBox(); };


            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem addToToolBoxMenuItem = new ToolStripMenuItem("&Add to tool box", bmp);
            addToToolBoxMenuItem.ImageTransparentColor = System.Drawing.Color.Black;
            addToToolBoxMenuItem.Click +=
                delegate { addSelectedControlToToolBox(); };
            contextMenu.Items.Add(addToToolBoxMenuItem);

            ToolStripMenuItem about = new ToolStripMenuItem("A&bout");
            about.Click += delegate { new AboutBox().ShowDialog(); };
            contextMenu.Items.Add(new ToolStripSeparator());
            contextMenu.Items.Add(about);
            treeView.BeforeSelect +=
                delegate(object sender, TreeViewCancelEventArgs e)
                {
                    addToToolBoxButton.Enabled = e.Node.Tag != null;
                    addToToolBoxMenuItem.Enabled = e.Node.Tag != null;
                };
            treeView.ContextMenuStrip = contextMenu;
            treeView.MouseDown +=
                delegate(object sender, MouseEventArgs e)
                {
                    if (e.Button == MouseButtons.Right)
                    {
                        TreeNode node = treeView.HitTest(e.Location).Node;
                        if (node != null)
                        {
                            treeView.SelectedNode = node;

                            if (treeView.SelectedNode.Tag != null)
                                contextMenu.Show(treeView, e.Location);
                        }
                    }
                };
            SelectionHandler selectionHandler = new SelectionHandler(treeView);
            treeView.KeyDown +=
                delegate(object sender, KeyEventArgs e)
                {
                    if (e.KeyData == Keys.Apps && treeView.SelectedNode.Tag != null)
                        contextMenu.Show(treeView, 0, 0);
                };
            bool reloading = false;
            System.Windows.Forms.ToolStripProgressBar progress = new ToolStripProgressBar();
            progress.Visible = false;
            progress.Height = 10;
            progress.ProgressBar.Height = 10;
            System.Windows.Forms.ToolStripLabel loading = new ToolStripLabel();
            loading.Text = "Loading...";
            loading.Visible = false;
            Thread theThread = null;

            Bitmap bmp1 = Resources.Refresh;
            ToolStripButton refreshButton = new ToolStripButton(bmp1);
            refreshButton.ToolTipText = "Refresh";
            refreshButton.ImageTransparentColor = System.Drawing.Color.Black;
            toolStrip.Items.Add(refreshButton);
            Action<bool> end =
                delegate(bool treeEnabled)
                {
                    reloading = false;
                    DoOnUIThread(delegate()
                                     {
                                         progress.Visible = false;
                                         loading.Visible = false;
                                         refreshButton.Enabled = true;
                                         treeView.Enabled = treeEnabled;
                                     });
                    selectionHandler.EndLoading();

                    theThread = null;
                };

            Action<string> reportMessage =
                delegate(string message)
                {
                    IVsOutputWindow outWindow = GetService(typeof(SVsOutputWindow)) as IVsOutputWindow;
                    if (outWindow != null)
                    {
                        Guid generalWindowGuid =
                            Microsoft.VisualStudio.VSConstants.GUID_OutWindowGeneralPane;
                        IVsOutputWindowPane windowPane;

                        outWindow.GetPane(ref generalWindowGuid, out windowPane);
                        windowPane.OutputString(message);
                    }
                };

            Action<Exception> reportException =
                delegate(Exception ex)
                {
                    reportMessage(string.Format("{0}\n\n{1}\n", ex.Message, ex.StackTrace));
                };
            Action<bool> reloadTree =
                delegate(bool forceRefresh)
                {
                    if (reloading)
                        return;


                    reloading = true;

                    EnvDTE.DTE dte = (EnvDTE.DTE)GetService(typeof(EnvDTE.DTE));
                    IDesignerHost designerHost = GetActiveDesigner(dte);
                    if (designerHost == null)
                    {
                        end(false);
                        return;
                    }

                    IToolboxService toolboxService = (IToolboxService)designerHost.GetService(
                                                                           typeof(IToolboxService));
                    Type rootBaseType = null;
                    try
                    {
                        if (designerHost.RootComponent == null)
                        {
                            end(false);
                            return;
                        }
                        rootBaseType = designerHost.RootComponent.GetType();
                    }
                    catch
                    {
                        end(false);
                        return;
                    }
                    if (!forceRefresh && _lastDesignedType != typeof(object))
                    {
                        end(_lastDesignedType == rootBaseType);
                        return;
                    }
                    _lastDesignedType = rootBaseType;
                    DoOnUIThread(delegate()
                                     {
                                         progress.Visible = true;
                                         loading.Visible = true;
                                         refreshButton.Enabled = false;
                                         progress.Value = 1;
                                         treeView.Nodes.Clear();
                                     });
                    selectionHandler.Reloadion();

                    theThread = new Thread(
                        delegate()
                        {
                            lock (reloadLock)
                            {
                                try
                                {
                                    VSLangProj.VSProject proj = (VSLangProj.VSProject)dte.ActiveDocument.ProjectItem.ContainingProject.Object;
                                    string projectOutputFileName =
                                        System.IO.Path.Combine(
                                            proj.Project.ConfigurationManager.
                                                ActiveConfiguration
                                                .Properties.Item("OutputPath").Value.
                                                ToString(),
                                            proj.Project.Properties.Item("OutputFileName").
                                                Value.
                                                ToString
                                                ());
                                    string fullProjectOutputFileName =
                                        System.IO.Path.IsPathRooted(projectOutputFileName)
                                            ?
                                        projectOutputFileName
                                            : System.IO.Path.Combine(
                                                  System.IO.Path.GetDirectoryName(
                                                      proj.Project.FileName),
                                                  projectOutputFileName);

                                    AddAssembly addAssembly =
                                        delegate(string assemblyPath, bool publicTypesOnly)
                                        {
                                            AppDomainSetup ads = new AppDomainSetup();
                                            ads.ApplicationBase = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                                            AppDomain appDomainForAssemblyLoading = AppDomain.CreateDomain(
                                                "AppDomainForAssemblyLoading", null, ads);

                                            System.Collections.Generic.SortedDictionary<string,
                                                System.Collections.Generic.SortedDictionary<string, TreeNode>> nodes =
                                                    new System.Collections.Generic.SortedDictionary<string,
                                                        System.Collections.Generic.SortedDictionary<string, TreeNode>>();
                                            AssemblyName assemblyName = null;

                                            try
                                            {
                                                assemblyName = ((AssemblyNameProxy)
                                                appDomainForAssemblyLoading.CreateInstanceAndUnwrap(
                                                    typeof(AssemblyNameProxy).Assembly.FullName,
                                                    typeof(AssemblyNameProxy).FullName)).GetAssemblyName(assemblyPath);
                                                ResolveEventHandler resolve =
                                                    delegate(object sender, ResolveEventArgs args)
                                                    {
                                                        return args.Name == Assembly.GetExecutingAssembly().FullName ?
                                                            Assembly.GetExecutingAssembly() : null;
                                                    };
                                                AppDomain.CurrentDomain.AssemblyResolve += resolve;
                                                AssemblyLoader assemblyLoader = (AssemblyLoader)
                                                    appDomainForAssemblyLoading.CreateInstanceFromAndUnwrap(
                                                    Assembly.GetExecutingAssembly().Location,
                                                    typeof(AssemblyLoaderClass).FullName);
                                                AppDomain.CurrentDomain.AssemblyResolve -= resolve;

                                                assemblyLoader.DoOnTypes(assemblyName, System.IO.Path.GetDirectoryName(fullProjectOutputFileName),
                                                    new AssemblyLoaderClientClass(
                                                    delegate(string typeNamespace, ToolboxItem item)
                                                    {
                                                        try
                                                        {
                                                            if (toolboxService.IsSupported(toolboxService.SerializeToolboxItem(item), designerHost))
                                                            {
                                                                item.Lock();
                                                                System.Windows.Forms.TreeNode node = new TreeNode();
                                                                DoOnUIThread(
                                                                    delegate()
                                                                    {
                                                                        node = new TreeNode(item.DisplayName);
                                                                        node.Tag = item;
                                                                        imageList.Add(item.TypeName, item.Bitmap);
                                                                        node.ImageIndex = imageList.GetImageIndexFor(item.TypeName);
                                                                        node.SelectedImageIndex = imageList.GetImageIndexFor(item.TypeName);
                                                                    });

                                                                System.Collections.Generic.SortedDictionary<string, TreeNode>
                                                                    componentNodes;
                                                                if (!nodes.TryGetValue(typeNamespace, out componentNodes))
                                                                {
                                                                    componentNodes =
                                                                        new System.Collections.Generic.SortedDictionary<
                                                                            string, TreeNode>();
                                                                    nodes.Add(typeNamespace, componentNodes);
                                                                }
                                                                componentNodes.Add(item.DisplayName, node);
                                                            }
                                                        }
                                                        catch (Exception e)
                                                        {
                                                            try
                                                            {
                                                                reportMessage(
                                                                    string.Format(
                                                                        "Firefly Toolbox error - Exception occured on load of " +
                                                                        item.TypeName +
                                                                        " the exception is:\n" + e));
                                                            }
                                                            catch { }
                                                        }
                                                    }), publicTypesOnly);
                                            }
                                            catch (Exception e)
                                            {
                                                try
                                                {
                                                    DoOnUIThread(
                                                        delegate()
                                                        {
                                                            System.Windows.Forms.TreeNode node = new TreeNode();
                                                            node.Text = "Error loading " + assemblyName + " - " +
                                                                        e.ToString();
                                                            treeView.Nodes.Add(node);
                                                        });

                                                    reportMessage(
                                                        string.Format(
                                                            "Firefly Toolbox error - Exception occured on load of " +
                                                            assemblyName.ToString() +
                                                            " the exception is:\n" + e));
                                                    ReflectionTypeLoadException le = e as ReflectionTypeLoadException;
                                                    if (le != null)
                                                        foreach (Exception ie in le.LoaderExceptions)
                                                        {
                                                            reportMessage(
                                                                string.Format(
                                                                    "loader exception exception exception is:\n" +
                                                                    ie));
                                                        }
                                                }
                                                catch { }
                                            }
                                            finally
                                            {
                                                AppDomain.Unload(appDomainForAssemblyLoading);
                                            }
                                            if (nodes.Count > 0)
                                            {
                                                DoOnUIThread(
                                                    delegate
                                                    {
                                                        treeView.BeginUpdate();
                                                        try
                                                        {
                                                            System.Windows.Forms.TreeNode assemblyNode =
                                                                new TreeNode(assemblyName.Name);
                                                            assemblyNode.ImageIndex = imageList.GetImageIndexFor("Assembly");
                                                            assemblyNode.SelectedImageIndex = imageList.GetImageIndexFor("Assembly");
                                                            treeView.Nodes.Add(assemblyNode);
                                                            foreach (System.Collections.Generic.KeyValuePair<string,
                                                                System.Collections.Generic.SortedDictionary
                                                                    <string, TreeNode>>
                                                                pair in nodes)
                                                            {
                                                                TreeNode namespaceNode = new TreeNode(pair.Key);
                                                                namespaceNode.ImageIndex = imageList.GetImageIndexFor("Namespace");
                                                                namespaceNode.SelectedImageIndex = imageList.GetImageIndexFor("Namespace");
                                                                assemblyNode.Nodes.Add(namespaceNode);
                                                                foreach (TreeNode n in pair.Value.Values)
                                                                {
                                                                    namespaceNode.Nodes.Add(n);
                                                                    selectionHandler.Loaded(n);
                                                                }
                                                                imageList.CommitToUi();
                                                            }
                                                        }
                                                        finally
                                                        {
                                                            treeView.EndUpdate();
                                                            if (treeView.SelectedNode != null)
                                                            {
                                                                treeView.Update();
                                                                treeView.Select();
                                                            }
                                                        }
                                                    });
                                            }
                                        };


                                    System.Collections.Generic.SortedDictionary<string, Cmd> addAssemblies =
                                        new System.Collections.Generic.SortedDictionary<string, Cmd>();

                                    if (System.IO.File.Exists(fullProjectOutputFileName))
                                    {
                                        addAssemblies[Path.GetFileName(fullProjectOutputFileName)] =
                                            delegate
                                            {
                                                addAssembly(fullProjectOutputFileName, false);
                                            };
                                    }
                                    foreach (
                                        VSLangProj.Reference reference in proj.References)
                                    {
                                        string path = reference.Path;
                                        addAssemblies[Path.GetFileName(path)] =
                                            delegate
                                            {
                                                addAssembly(path, true);
                                            };
                                    }


                                    DoOnUIThread(
                                        delegate()
                                        {
                                            progress.ProgressBar.Maximum = addAssemblies.Count + 1;
                                        });
                                    foreach (Cmd cmd in addAssemblies.Values)
                                    {
                                        cmd();
                                        DoOnUIThread(
                                            delegate()
                                            {
                                                try
                                                {
                                                    progress.ProgressBar.Value++;
                                                }
                                                catch
                                                {
                                                }
                                            });
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if (!(ex is ThreadAbortException))
                                    {
                                        reportException(ex);
                                    }
                                }
                                finally
                                {
                                    end(true);
                                }
                            }
                        });
                    theThread.Start();
                };

            loading.Click += delegate
                                 {
                                     reloading = false;
                                     theThread.Abort();
                                     end(true);
                                 };

            refreshButton.Click +=
                delegate { reloadTree(true); };


            toolStrip.Items.Add(progress);
            toolStrip.Items.Add(loading);


            mySelectionListener listener = new mySelectionListener(treeView, delegate() { reloadTree(false); });

            control.Load +=
                delegate
                {
                    IVsMonitorSelection monitor = (IVsMonitorSelection)GetService(typeof(IVsMonitorSelection));
                    monitor.AdviseSelectionEvents(listener, out cookie);
                    reloadTree(false);
                };
            control.Disposed +=
                delegate
                {
                    IVsMonitorSelection monitor = (IVsMonitorSelection)GetService(typeof(IVsMonitorSelection));
                    monitor.UnadviseSelectionEvents(cookie);
                };
            treeView.ItemDrag +=
                delegate(object sender, ItemDragEventArgs e)
                {
                    IToolboxService toolboxService = (IToolboxService)this.GetService(typeof(IToolboxService));
                    System.Windows.Forms.TreeNode node = (System.Windows.Forms.TreeNode)e.Item;
                    if (node.Tag == null) return;
                    System.Windows.Forms.DataObject dataObject = toolboxService.SerializeToolboxItem(
                                (ToolboxItem)node.Tag) as System.Windows.Forms.DataObject;
                    treeView.DoDragDrop(dataObject,
                                System.Windows.Forms.DragDropEffects.All);
                };
        }

        object reloadLock = new object();
        void DoOnUIThread(Cmd cmd)
        {
            Cmd command = delegate { try { cmd(); } catch { } };
            this.control.Invoke(command);
        }

        delegate void AddAssembly(string assemblyPath, bool publicTypesOnly);
        private static ToolboxItem CreateToolboxItem(Type type, AssemblyName name)
        {
            ToolboxItemAttribute attribute1 = (ToolboxItemAttribute)
                TypeDescriptor.GetAttributes(type)[typeof(ToolboxItemAttribute)];
            ToolboxItem item1 = null;
            if (!attribute1.IsDefaultAttribute())
            {
                Type type1 = attribute1.ToolboxItemType;
                if (type1 != null)
                {
                    item1 = CreateToolboxItemInstance(type1, type);
                    if (item1 != null)
                    {
                        if (name != null)
                        {
                            item1.AssemblyName = name;
                        }
                    }
                }
            }
            if (((item1 == null) && (attribute1 != null)) && !attribute1.Equals(ToolboxItemAttribute.None))
            {
                item1 = new ToolboxItem(type);
                if (name != null)
                {
                    item1.AssemblyName = name;
                }
            }
            return item1;
        }
        private static ToolboxItem CreateToolboxItemInstance(Type toolboxItemType, Type toolType)
        {
            ToolboxItem item1 = null;
            ConstructorInfo info1 = toolboxItemType.GetConstructor(new Type[] { typeof(Type) });
            if (info1 != null)
            {
                return (ToolboxItem)info1.Invoke(new object[] { toolType });
            }
            info1 = toolboxItemType.GetConstructor(new Type[0]);
            if (info1 != null)
            {
                item1 = (ToolboxItem)info1.Invoke(new object[0]);
                item1.Initialize(toolType);
            }
            return item1;
        }
        interface AssemblyLoader
        {
            void DoOnTypes(AssemblyName assemblyName, string path, AssemblyLoaderClient client, bool publicTypesOnly);
        }
        interface AssemblyLoaderClient
        {
            void DoOnToolboxItemForType(string typeNamespace, ToolboxItem toolboxItem, byte[] bitmapData);
        }
        class AssemblyLoaderClass : MarshalByRefObject, AssemblyLoader
        {
            public void DoOnTypes(AssemblyName assemblyName, string path, AssemblyLoaderClient client, bool publicTypesOnly)
            {
                ResolveEventHandler resolve =
                    delegate(object sender, ResolveEventArgs args)
                    {
                        var an = Path.Combine(path, args.Name.Split(',')[0] + ".dll");
                        return System.IO.File.Exists(an) ? Assembly.LoadFile(an) : null;
                    };
                AppDomain.CurrentDomain.AssemblyResolve += resolve;
                Assembly assembly = Assembly.Load(assemblyName);

                foreach (Type type in assembly.GetTypes())
                {
                    if (type.IsAbstract) continue;
                    if (type.IsNestedPrivate) continue;
                    if (publicTypesOnly && !type.IsVisible) continue;
                    if (!typeof(System.ComponentModel.IComponent).IsAssignableFrom(type))
                        continue;
                    ToolboxItem item = CreateToolboxItem(type, assembly.GetName());
                    if (item == null) continue;
                    client.DoOnToolboxItemForType(type.Namespace, item,
                        (byte[])TypeDescriptor.GetConverter(item.Bitmap).ConvertTo(item.Bitmap, typeof(byte[])));
                }
                AppDomain.CurrentDomain.AssemblyResolve -= resolve;
            }
        }

        delegate void DoOnToolboxItemForType(string typeNamespace, ToolboxItem toolboxItem);
        class AssemblyLoaderClientClass : MarshalByRefObject, AssemblyLoaderClient
        {
            DoOnToolboxItemForType _doOnToolboxItemForType;


            public AssemblyLoaderClientClass(DoOnToolboxItemForType doOnToolboxItemForType)
            {
                _doOnToolboxItemForType = doOnToolboxItemForType;
            }

            public void DoOnToolboxItemForType(string typeNamespace, ToolboxItem toolboxItem, byte[] bitmapData)
            {
                using (MemoryStream stream = new MemoryStream(bitmapData))
                {
                    Bitmap theBitmap = (Bitmap)Bitmap.FromStream(stream);
                    theBitmap.MakeTransparent();
                    toolboxItem.Bitmap =
                        (Bitmap)TypeDescriptor.GetConverter(typeof(Bitmap)).ConvertFrom(bitmapData);
                    _doOnToolboxItemForType(typeNamespace, toolboxItem);
                }
            }
        }

        class myDataObject : IDataObject
        {
            IDataObject _data;
            public myDataObject(IDataObject data)
            {
                _data = data;
            }
            public object GetData(Type format)
            {
                return _data.GetData(format);
            }

            public object GetData(string format)
            {
                return _data.GetData(format);
            }

            public object GetData(string format, bool autoConvert)
            {
                return _data.GetData(format, autoConvert);
            }

            public bool GetDataPresent(Type format)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("GetDataPresent {0}", format));
                return _data.GetDataPresent(format);
            }

            public bool GetDataPresent(string format)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("GetDataPresent {0}", format));
                return _data.GetDataPresent(format);
            }

            public bool GetDataPresent(string format, bool autoConvert)
            {
                System.Diagnostics.Debug.WriteLine(string.Format("GetDataPresent {0}", format));
                return _data.GetDataPresent(format, autoConvert);
            }

            public string[] GetFormats()
            {
                System.Diagnostics.Debug.WriteLine("GetForamts");
                return _data.GetFormats();
            }

            public string[] GetFormats(bool autoConvert)
            {
                System.Diagnostics.Debug.WriteLine("GetForamts");
                return _data.GetFormats(autoConvert);
            }

            public void SetData(object data)
            {
                _data.SetData(data);
            }

            public void SetData(Type format, object data)
            {
                _data.SetData(format, data);
            }

            public void SetData(string format, object data)
            {
                _data.SetData(format, data);
            }

            public void SetData(string format, bool autoConvert, object data)
            {
                _data.SetData(format, autoConvert, data);
            }
        }
        IDesignerHost GetActiveDesigner(EnvDTE.DTE DTE)
        {
            if (DTE.ActiveDocument != null)
                foreach (EnvDTE.Window W in DTE.ActiveDocument.Windows)
                {
                    IDesignerHost Host = W.Object as IDesignerHost;
                    if (Host != null)
                        return Host;
                }
            return null;
        }
        delegate void Cmd();
        class mySelectionListener : IVsSelectionEvents
        {
            System.Windows.Forms.Control _control;
            Cmd _activeFormChanged;
            public mySelectionListener(System.Windows.Forms.Control control,
                Cmd activeFormChanged)
            {
                _control = control;
                _activeFormChanged = activeFormChanged;
            }
            public int OnCmdUIContextChanged(uint dwCmdUICookie, int fActive)
            {
                return 0;
            }

            public int OnElementValueChanged(uint elementid, object varValueOld, object varValueNew)
            {
                return 0;
            }

            public int OnSelectionChanged(IVsHierarchy pHierOld, uint itemidOld, IVsMultiItemSelect pMISOld, ISelectionContainer pSCOld, IVsHierarchy pHierNew, uint itemidNew, IVsMultiItemSelect pMISNew, ISelectionContainer pSCNew)
            {
                _activeFormChanged();
                return 0;
            }
        }

        /// <summary>
        /// This property returns the handle to the user control that should
        /// be hosted in the Tool Window.
        /// </summary>
        override public IWin32Window Window
        {
            get
            {
                return (IWin32Window)control;
            }
        }
    }
}
