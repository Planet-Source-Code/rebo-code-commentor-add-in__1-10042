Created by: rebo@geocities.com
Last update: 25 july 2000

This Add-in is written in VB6 and sp4.
You have to manualy register the add-in dll.
First copy the RBVBCommentor.dll to the \WINDOWS\SYSTEM32 or \WINNT\SYSTEM32 directory.
In the SYSTEM32 directory type REGSVR32 RBVBCommentor.dll.
If the registration is succeeded, start Visual Basic and activate the add-in.
In some cases the add-in is not shown in the add-in manager.
To make it visible, go to the \WINDOWS or \WINNT directory and edit the VBADDIN.INI.
Add the following key; RBVBCommentor.Connector=0.
Restart Visual Basic and activate the Add-in.