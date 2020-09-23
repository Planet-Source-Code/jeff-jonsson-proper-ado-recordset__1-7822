<div align="center">

## Proper ADO Recordset


</div>

### Description

To remedy difficiencies in ADO, namely single criteria in the .FIND method.

We encapsulated an ADODB.Recordset within a vb6 class module, and created four methods (FindFirst, FindLast, FindNext, FindPrevious) which allow for more that one criteria.
 
### More Info
 
Use is almost exactly that of a standard ADODB.Recordset.

Just drop the class module into a project, and rename all:

Dim aRecSet as ADODB.Recordset

to:

Dim aRecSet as ProperADORecordset

then CTRL-F5 to find and replace all Open/Close methods with OpenIt/CloseIt.

Any new properties/methods simply pass arguments/return values between the encapsulated ADODB.Recordset and the created class module.

Returns less headaches.

First, not all properties/methods have been emulated. Easy enough to do though. Left as an exercise.

Second, due to vb limitations, Open/Close not supported, but renamed as OpenIt/CloseIt.


<span>             |<span>
---                |---
**Submitted On**   |2000-05-03 14:59:34
**By**             |[Jeff Jonsson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeff-jonsson.md)
**Level**          |Advanced
**User Rating**    |4.7 (33 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD5465532000\.zip](https://github.com/Planet-Source-Code/jeff-jonsson-proper-ado-recordset__1-7822/archive/master.zip)








