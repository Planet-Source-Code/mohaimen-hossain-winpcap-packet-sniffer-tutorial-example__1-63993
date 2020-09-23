Attribute VB_Name = "VBPCAP"
'************************************************************************
' Visual Basic Packet Capture
'A packet capture engine for Visual Basic* (c)
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation; either
'version 2.1 of the License, or (at your option) any later version.
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

'*Visual Basic is a registered trademark of Micorsoft corporation */
'*************************************************************************


Public Declare Function VBPcapInit Lib "VBPCAP.DLL" () As Long
Public Declare Function VBPcapTerminate Lib "VBPCAP.DLL" () As Long
Public Declare Function vpBegin Lib "VBPCAP.DLL" (ByVal timeout As Long) As Long
Public Declare Function vpEnd Lib "VBPCAP.DLL" () As Long
Public Declare Function vpGetAdapterInfo Lib "VBPCAP.DLL" (ByVal ID As Integer, ad As AdINFO) As Long
Public Declare Function vpCaptureMem Lib "VBPCAP.DLL" (ByRef buffer() As Byte, hed As PacketHeader) As Long
Public Declare Function vpCaptureMemSafe Lib "VBPCAP.DLL" (ByRef buffer() As Byte) As Long
Public Declare Function vpSetCurrentAdapter Lib "VBPCAP.DLL" (ByVal ID As Integer) As Long
Public Declare Function vpGetCurrentAdapter Lib "VBPCAP.DLL" () As Long
Public Declare Function vpGetErrorDescription Lib "VBPCAP.DLL" () As String
Public Declare Function vpSetKernelBuffSize Lib "VBPCAP.DLL" (ByVal bSize As Long) As Long
Public Declare Function SetCaptureParams Lib "VBPCAP.DLL" (ByVal ID As Long, data As Variant) As Long
Public Declare Function vpCapture Lib "VBPCAP.DLL" (ByRef buffer() As Byte, hed As PacketHeader) As Long
Public Declare Function vpSetParam Lib "VBPCAP.DLL" (ByVal param As VBPCAPPARAMS, value As Variant) As Long
Public Declare Function vpGetAdapterInfoVB5 Lib "VBPCAP.DLL" (ByVal ID As Integer, name As String, desc As String) As Long
Public Declare Function vpCaptureDiskSafe Lib "VBPCAP.DLL" (hed As PacketHeader) As Long

'---------USE THIS ENUM WITH THE CORRESPONDING PARAMETERS------------
Public Enum VBPCAPPARAMS
PRM_SNAPLEN = 0
PRM_MODE = 1 ' accept MODE Enum
PRM_FILENAME = 2
PRM_KERNELBUFFSIZE = 6 'accept KERNELBUFFSIZE enum
PRM_DUMPTYPE = 8 'accept DUMPTYPE enum
End Enum

Public Enum DUMPTYPE
MEM = 10
MEMSAFE = 100
DISK_SAFE = 1
End Enum

Public Enum MODE
CAPTURE_PROMISCUOUS = 1
CAPTURE_LOCAL = 0
End Enum

Public Enum KERNELBUFFSIZE
A_1_MegaBytes = 1024
B_2_MegaBytes = 2048
C_3_MegaBytes = 3072
D_4_MegaBytes = 4096
E_5_MegaBytes = 5120
End Enum


Public Enum VBCAPI_ERRORS
VBAPI_ERROR_SUCCESS = 0
VBAPI_ERROR_FAILURE = -1
End Enum


Type AdINFO
name As String * 255
Description As String * 255
Loopback As Long
End Type


Type PacketHeader
caplen As Long
len As Long
h As Long
m As Long
s As Long
us As Long
End Type



