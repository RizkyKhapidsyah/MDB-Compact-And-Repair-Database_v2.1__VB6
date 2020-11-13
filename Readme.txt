CompMDB v2.1

Important:  The two EXE files must be in the same 
            directory.  The main application will
            call _DelTemp.exe with you terminate the 
            CompMDb.exe application.  Its purpose is
            to remove all temporary folders created
            by the calling application.  Source code
            is included with this utility.

This program will allow the user to select a MDB file
to compact.  The size of the file is captured and a
calculation of twice that size is made to determine
the amount of free space required to compact the
database.  Half that amount is used for a backup copy
of the original database and the other half is for
the compacted database.  if there is not enough space,
the user is prompted to select another path in which
to perform this operation or leave the application.
After the database is compacted, the original is deleted
and the new version is moved back into the place of the
original.

This program now recognizes drives greater than 2gb and
you can use command line parameters to point to your
favorite database.

I am referencing Microsoft DAO 3.6 Object Library.
DLL name is DAO360.dll located in 
C:\Program Files\Common\Microsoft Shared\DAO folder.
If you have a problem, then check your references first.
In VB design mode, select Project, References.  This 
has the repair built into the compact.  It is no longer
a separate command.

I am using two components:
  Microsoft Common Dialog Control 6.0 (sp3)
  Microsoft Windows Common Controls-2 6.0 (sp3)
    (In this, I am accessing the Animation control)

You will note that I am using a resource file to store the
AVI file. 

Read the documentation in the code to understand what
I am doing.

-----------------------------------------------------------------
Written by Kenneth Ives                    kenaso@home.com

All of my routines have been compiled with VB6 Service Pack 3.
There are several locations on the web to obtain these
modules.

Whenever I use someone else's code, I will give them credit.  
This is my way of saying thank you for your efforts.  I would
appreciate the same consideration.

Read all of the documentation within this program.  It is very
informative.  Also, if you learn to document properly now, you
will not be scratching your head next year trying to figure out
exactly what you were programming today.  Been there, done that.

This software is FREEWARE. You may use it as you see fit for 
your own projects but you may not re-sell the original or the 
source code. If you redistribute it you must include this 
disclaimer and all original copyright notices. 

No warranty express or implied, is given as to the use of this
program. Use at your own risk.

If you have any suggestions or questions, I'd be happy to
hear from you.
-----------------------------------------------------------------

 
 