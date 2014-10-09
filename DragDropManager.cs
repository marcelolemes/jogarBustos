using System;
using System.Windows.Forms;
using System.Diagnostics;

// Written by: Alex Farber
//
// alexm@cmt.co.il


namespace DropFiles1
{
    /// <summary>
    /// Interface which should be implemented by parent form
    /// to open files dropped from Windows Explorer.
    /// </summary>
    public interface IDropFileTarget
    {
        void OpenFiles(System.Array a);
    }


	/// <summary>
	/// DragDropManager class allows to open files dropped from 
	/// Windows Explorer in Windows Form application.
	/// 
	/// Using:
	/// 1) Derive parent form from IDropFileTarget interface:
	/// 
	///          public class Form1 : System.Windows.Forms.Form, IDropFileTarget
	///     
	/// 2) Implement IDropFileTarget interface in parent form:
	/// 
    ///          public void OpenFiles(Array a)
    ///          {
    ///              // open files from array here
    ///          }
    ///         
    ///  3) Add member of this class to parent form:
    ///  
    ///          private DragDropManager m_DragDropManager;
	/// 
	///  4) Initialize class instance in parent form Load event:
	///  
    ///          private void Form1_Load(object sender, System.EventArgs e)
    ///          {
    ///              m_DragDropManager = new DragDropManager();
    ///              m_DragDropManager.Parent = this;
    ///          }
	/// 
	/// </summary>
	public class DragDropManager
	{
        private Form m_parent;          // reference to owner form

        // delegate used in asynchronous call to parent form:
        private delegate void DelegateOpenFiles(Array a);           // type
        private DelegateOpenFiles m_DelegateOpenFiles;              // instance

        public DragDropManager()
		{
		}

        /// <summary>
        /// Set reference to parent form and make initialization.
        /// </summary>
        public Form Parent
        {
            set
            {
                m_parent = value;               // keep reference to parent form

                // Check if parent form implements IDropFileTarget interface
                if ( ! ( m_parent is IDropFileTarget ) )
                {
                    throw new Exception(
                        "DragDropManager: Parent form doesn't implement IDropFileTarget interface");
                }

                // create delegate used for asynchronous call
                m_DelegateOpenFiles = new DelegateOpenFiles(((IDropFileTarget)m_parent).OpenFiles);

                // ensure that parent form allows dropping
                m_parent.AllowDrop = true;

                // subscribe to parent form's drag-drop events
                m_parent.DragEnter += new System.Windows.Forms.DragEventHandler(this.OnDragEnter);
                m_parent.DragDrop += new System.Windows.Forms.DragEventHandler(this.OnDragDrop);
            }
        }

        /// <summary>
        /// Handle parent form DragEnter event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnDragEnter(object sender, System.Windows.Forms.DragEventArgs e)
        {
            // If file is dragged, show cursor "Drop allowed"
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) 
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        /// <summary>
        /// Handle parent form DragDrop event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnDragDrop(object sender, System.Windows.Forms.DragEventArgs e)
        {
            try
            {
                // When file(s) are dragged from Explorer to the form, IDataObject
                // contains array of file names. If one file is dragged,
                // array contains one element.
                Array a = (Array)e.Data.GetData(DataFormats.FileDrop);

                if ( a != null )
                {
                    // Call parent's OpenFiles asynchronously.
                    // Explorer instance from which file is dropped is not responding
                    // all the time when DragDrop handler is active, so we need to return
                    // immidiately (especially if OpenFiles shows MessageBox).

                    m_parent.BeginInvoke(m_DelegateOpenFiles, new Object[] {a});

                    m_parent.Activate();        // in the case Explorer overlaps parent form
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Error in DragDropManager.OnDragDrop function: " + ex.Message);

                // don't show MessageBox here - Explorer is waiting !
            }
        }

	}
}
