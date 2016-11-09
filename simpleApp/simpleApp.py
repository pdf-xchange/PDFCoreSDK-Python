import os
import sys
import comtypes
from ctypes import c_ulong, byref
from comtypes.client import GetModule, CreateObject
path = os.environ['PXCE_BIN32D_PATH']
path += "\\PDFXCoreAPI.x86.dll"
# generate wrapper code for the type library, this needs
# to be done only once (but also each time the IDL file changes)
GetModule(path);

import comtypes.gen.PDFXCoreAPI as PDFXCoreAPI

g_Inst = CreateObject("PDFXCoreAPI.PXC_Inst", None, None, PDFXCoreAPI.IPXC_Inst)
g_Inst.Init(sKey = "");
try:
    pDoc = g_Inst.NewDocument();
    pDoc.Props.SpecVersion = 0x10007;

    pr = PDFXCoreAPI.PXC_Rect(0.0,0.0, 8.5 * 72.0, 11.0 * 72.0);
    for i in range(0, 6):
        pDoc.Pages._IPXC_Pages__com_AddEmptyPages(c_ulong(-1), 1, byref(pr), None, None)
        #pDoc.Pages.AddEmptyPages(c_ulong(-1), 1, byref(pr), None)

    pDoc.WriteToFile(sys.argv[0] + u".pdf")
except comtypes.COMError as e:
    print(e.text)
pDoc.Close()
g_Inst.Finalize()
