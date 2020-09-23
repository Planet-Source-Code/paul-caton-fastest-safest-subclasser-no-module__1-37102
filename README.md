<div align="center">

## Fastest, safest subclasser, no module\!


</div>

### Description

***

Update: See my new submission here...

http://www.exhedra.com/vb/scripts/ShowCode.asp?txtCodeId=42918&lngWId=1

If you do want the original zip then email me at Paul_Caton@hotmail.com

***

cSuperClass.cls is i believe the fastest, safest compiled in window subclasser around. Speed: The WndProc is executed entirely in run-time dynamically generated machine code. The class only calls back on messages that you choose. Safety: So far I've not been able to crash the IDE by pressing the end button or with the End statement. Flexible: The programmer can choose between filtered mode (fastest) and all messages mode. In filtered mode the user decides which windows messages they're interested in and can individually specify whether the message is to callback after default processing or before. Before mode additionally allows the programmer to specify whether or not default processing is to be performed subsequently.

No module: AFAIK this is the only subclasser ever to eschew the use of a module. So how do I get the address of the WndProc routine? Simple, the dynamically generated machine code lives in a byte array; you can get its address with the undocumented VarPtr function. The real magic in cSuperClass.cls is getting from the WndProc to the callback interface routine using ObjPtr against the owning Form/UserControl, see the assembler .asm model file included in the zip. Speaking of which... it may well be the case that my assembler is sub-optimal. Any experts out there willing to take a look? I thought I had a nifty/dirty stack trick working for a while but it didn't pan out. Should work with VB5 if VarPtr & ObjPtr were in that release? Sample project included. Regards.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-07-06 11:26:52
**By**             |[Paul Caton](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-caton.md)
**Level**          |Advanced
**User Rating**    |5.0 (422 globes from 85 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Fastest,\_s1085867202002\.zip](https://github.com/Planet-Source-Code/paul-caton-fastest-safest-subclasser-no-module__1-37102/archive/master.zip)








