using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace Firefly.Box.VSIntegration
{
    class SelectionHandler
    {
        string[] _selectedStack = new string[] { };
        bool _reloading = false;
        public SelectionHandler(System.Windows.Forms.TreeView tv)
        {
            tv.AfterSelect += delegate(object sender, TreeViewEventArgs e)
                                {
                                    if (_reloading)
                                        return;
                                    if (tv.SelectedNode == null)
                                        _selectedStack = null;
                                    else
                                        _selectedStack = CreateStack(tv.SelectedNode);
                                };
        }

        public void Reloadion()
        {
            _reloading = true;
        }

        string[] CreateStack(TreeNode tn)
        {
            List<string> l = new List<string>();
            while (tn != null)
            {
                l.Insert(0, tn.Text);
                tn = tn.Parent;
            }
            return l.ToArray();
        }


        public void Loaded(TreeNode n)
        {
            if (_selectedStack == null)
                return;
            string[] ar = CreateStack(n);
            if (ar.Length != _selectedStack.Length)
                return;
            for (int i = 0; i < ar.Length; i++)
            {
                if (ar[i] != _selectedStack[i])
                    return;
            }
            TreeNode parent = n.Parent;
            while (parent != null)
            {
                parent.Expand();
                parent = parent.Parent;
            }
            n.TreeView.SelectedNode = n;
        }

        public void EndLoading()
        {
            _reloading = false;
        }
    }
}
