<div align="center">

## Detecting if connected to LAN or WAN


</div>

### Description

One line of code will tell you if you are connected to LAN or WAN(Internet).
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Khursheed\_Siddiqui](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/khursheed-siddiqui.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/khursheed-siddiqui-detecting-if-connected-to-lan-or-wan__1-37004/archive/master.zip)





### Source Code

Private Declare Function IsNetworkAlive Lib "Sensapi.dll" (LPDFlags As Long) As Long <BR>
Private Const NETWORK_ALIVE_LAN = &H1 'net card connection<BR>
Private Const NETWORK_ALIVE_WAN = &H2 'RAS connection<BR>
Private Const NETWORK_ALIVE_AOL = &H4 'AOL<BR>
Private Sub Form_Load()<BR>
 Dim tmp As Long<BR>
 Dim ConnectionType As String<BR>
 If IsNetworkAlive(tmp) = NETWORK_ALIVE_LAN Then<BR>
 ConnectionType = "LAN"<BR>
 ElseIf IsNetworkAlive(tmp) = NETWORK_ALIVE_WAN Then<BR>
 ConnectionType = "WAN"<BR>
 ElseIf IsNetworkAlive(tmp) = NETWORK_ALIVE_AOL Then<BR>
 ConnectionType = "AOL"<BR>
 Else<BR>
 ConnectionType = "Could not Determine."<BR>
 End If<BR>
 Print<BR>
 Print "Your connection type is: " & ConnectionType<BR>
End Sub

