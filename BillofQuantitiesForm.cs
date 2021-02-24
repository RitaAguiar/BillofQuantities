#region Namespaces

using System;
using System.Windows.Forms;
using Autodesk.Revit.UI;

#endregion // Namespaces

namespace BillofQuantities
{
    public partial class BillofQuantitiesForm : Form
    {
        // In this sample, the dialog owns the handler and the event objects,
        // but it is not a requirement. They may as well be static properties
        // of the application.

        private RequestHandler m_Handler;
        private ExternalEvent m_ExEvent;

        string folderPath = null;

        //Dialog instantiation
        public BillofQuantitiesForm(ExternalEvent exEvent, RequestHandler handler)
        {
            InitializeComponent();

            m_Handler = handler;
            m_ExEvent = exEvent;
        }

        #region Form Items

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Folder path to save the report";
            if (fbd.ShowDialog() == DialogResult.OK) textBox1.Text = fbd.SelectedPath;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //save InputData static variables
            InputData.folderPath = textBox1.Text;
            InputData.instancesSheet = checkBox1.Checked;
            InputData.elementTypesSheet = checkBox2.Checked;
            InputData.billofQuantitiesSheet = checkBox3.Checked;

            if (textBox1.Text == null || textBox1.Text == "")
            {
                InputData.folderPath = folderPath = "C://Users//" + Environment.UserName + "//Documents";
            }

            //CALLS MAIN METHOD
            MakeRequest(RequestId.CreateBillofQuantities);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        //Exit - closing the dialog
        private void buttonExit_Click_1(object sender, EventArgs e)
        {
            Close();
        }

        #endregion //Form Items

        #region Form Events

        // Form closed event handler
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            // we own both the event and the handler
            // we should dispose it before we are closed
            m_ExEvent.Dispose();
            m_ExEvent = null;
            m_Handler = null;

            // do not forget to call the base class
            base.OnFormClosed(e);
        }

        //   Control enabler / disabler 
        private void EnableCommands(bool status)
        {
            foreach (System.Windows.Forms.Control ctrl in Controls)
            {
                ctrl.Enabled = status;
            }
            if (!status)
            {
                this.buttonExit.Enabled = true;
            }
        }

        //A private helper method to make a request
        //and put the dialog to sleep at the same time.

        //    It is expected that the process which executes the request 
        //   (the Idling helper in this particular case) will also
        //   wake the dialog up after finishing the execution.

        private void MakeRequest(RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
            DozeOff();
        }

        //DozeOff -> disable all controls (but the Exit button)
        private void DozeOff()
        {
            EnableCommands(false);
        }

        //WakeUp -> enable all controls
        public void WakeUp()
        {
            EnableCommands(true);
        }

        //Exit - closing the dialog
        private void buttonExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        #endregion //Form Events


    }
}
