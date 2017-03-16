using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Firefly.Box.VSIntegration
{
    class ImageListManager
    {
        System.Windows.Forms.ImageList _il;
        Dictionary<string, int> _keyIndex = new Dictionary<string, int>();
        int _lastImage = -1;
        public void SetToTreeView(System.Windows.Forms.TreeView tv)
        {
            tv.ImageList = _il;
        }

        List<Image> _waitingImages = new List<Image>();
        public ImageListManager()
        {
            
            _il = new ImageList();
        }

        public void Add(string key, Icon icon)
        {
            if (!_keyIndex.ContainsKey(key))
            {
                _keyIndex.Add(key, ++_lastImage);
                _il.Images.Add(icon);
            }
        }
        public void Add(string key,Image image)
        {
            if (!_keyIndex.ContainsKey(key))
            {
                _keyIndex.Add(key, ++_lastImage);
                _waitingImages.Add(image);
            }
        }

        public int GetImageIndexFor(string key)
        {
            return _keyIndex[key];
        }

        public void CommitToUi()
        {
            _il.Images.AddRange(_waitingImages.ToArray());
            _waitingImages = new List<Image>();
        }
    }
}
