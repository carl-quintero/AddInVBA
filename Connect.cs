using Extensibility;
using System;
using System.Runtime.InteropServices;

namespace MyAddInVBA
{
   [ProgId("AddInVBA.Connect"), Guid("B68B74CD-71AB-49D2-8B3E-3D3FAAA1935C"), ComVisible(true)]
   public class Connect : IDTExtensibility2
   {
      public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
      {
      }

      public void OnStartupComplete(ref Array custom)
      {
      }

      public void OnAddInsUpdate(ref Array custom)
      {
      }

      public void OnBeginShutdown(ref Array custom)
      {
      }

      public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
      {
      }

   }
}
