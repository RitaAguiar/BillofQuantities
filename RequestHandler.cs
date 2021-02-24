#region Copyright
//
// (C) Copyright 2003-2021 by Autodesk, Inc. All rights reserved.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.

//
// AUTODESK PROVIDES THIS PROGRAM 'AS IS' AND WITH ALL ITS FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable. 
#endregion Copyright

#region namespaces
using Autodesk.Revit.UI;
#endregion // namespaces

namespace BillofQuantities
{
    public partial class RequestHandler : IExternalEventHandler
    {
        // The value of the latest request made by the modeless form 
        private Request m_request = new Request();

        // A public property to access the current request value
        public Request Request
        {
            get { return m_request; }
        }

        //  A method to identify this External Event Handler
        public string GetName()
        {
            return "Revit 2021 External Event";
        }

        //The top method of the event handler

        //   This is called by Revit after the corresponding
        //   external event was raised (by the modeless form)
        //   and Revit reached the time at which it could call
        //   the event's handler (i.e. this object)

        public void Execute(UIApplication uiapp)
        {
            //Creates new instance everytime Excute is called by the IExternalEventHandler
            var instance = new RevitUtils();

            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        {
                            return;  // no request at this time -> we can leave immediately
                        }
                    case RequestId.CreateBillofQuantities:
                        {
                            instance.CreateBillOfQuantities(uiapp);
                            break;
                        }
                    default:
                        {
                            // some kind of a warning here should
                            // notify us about an unexpected request 
                            break;
                        }
                }
            }
            finally
            {
                Application.thisApp.WakeFormUp();
            }

            return;
        }
    }
}

