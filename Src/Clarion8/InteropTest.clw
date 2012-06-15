
  PROGRAM

  MAP
SetupOLE    PROCEDURE(SIGNED pFeq, STRING pOLEServer), SIGNED
  END

PositionGroup                 GROUP,TYPE         !Control coordinates
XPos                            SIGNED           !Horizontal coordinate
YPos                            SIGNED           !Vertical coordinate
Width                           UNSIGNED         !Width
Height                          UNSIGNED         !Height
                              END

Window    WINDOW('Caption'),AT(,,362,213),GRAY,IMM,SYSTEM,FONT('Segoe UI',,,, |
            CHARSET:DEFAULT)
            BUTTON('Exit'),AT(329,198),USE(?BUTTON1)
            PROMPT('Clarion Window with .NET COM object...'),AT(14,6),USE(?PROMPT1)
            REGION,AT(23,44,299,137),USE(?Region)
          END
oleFeq SIGNED 
  CODE
  
  Open(Window)
  ACCEPT
    CASE Event()
    OF EVENT:OpenWindow
      ! Uncomment these two lines to see it working with the Internet Explorer control:
      ! oleFeq = SetupOLE(?Region, 'Shell.Explorer.2')
      ! oleFeq{'Navigate2("http://fushnisoft.com")'}
      
      ! This will only work if the UCTest.dll is registered
      ! oleFeq = SetupOLE(?Region, 'UCTest.Test')
      
      ! However, this line should work to create the OLE object 
      oleFeq = SetupOLE(?Region, '{{B6F3818A-E877-388B-A783-0D18D16BA2F4}')
      ! If the UCTest.dll had been registered this would also work as well as showing the UCTest UserControl on the clarion window.
      ! Without registration the manifest redirection does work but only enough to allow calling methods in the class
      ! For some reason it doesn't put the object in the OLE container
      
      Stop('Calling GetDateTime method from UCTest.Test class<13,10,13,10>GetDateTime=' & oleFeq{'GetDateTime()'})
      
      
      ! Register the UCTest.dll with something like the following:
      ! c:\Windows\Microsoft.NET\Framework\v2.0.50727>regasm c:\Dev\ClarionRegFreeCOM\Bin\UCTest.dll /tlb /codebase
      ! 
      ! Unregister it like this:
      ! c:\Windows\Microsoft.NET\Framework\v2.0.50727>regasm c:\Dev\ClarionRegFreeCOM\Bin\UCTest.dll /u
      
    OF EVENT:Accepted
      IF Field() = ?BUTTON1
        Post(EVENT:CloseWindow)
      END
    END
  END
  
SetupOLE    PROCEDURE(SIGNED pFeq, STRING pOLEServer) !,SIGNED
pos           Like(PositionGroup)
newFeq        SIGNED
  CODE
  
  ! Remove the placeholder control
  GetPosition(pFeq, pos.XPos, pos.YPos, pos.Width, pos.Height)
  Destroy(pFeq)
  ! Now create a new control in the same place
  newFeq = Create(0, CREATE:OLE) 
  IF newFeq = 0
    Stop('CREATE:OLE failed')
    RETURN newFeq
  END

  SetPosition(newFeq, pos.XPos, pos.YPos, pos.Width, pos.Height)
  UnHide(newFeq)

  newFeq{PROP:ReportException} = 1
  newFeq{PROP:Create} = pOleServer
  newFeq{PROP:Compatibility} = 20h                                            ! 32 bit
  newFeq{PROP:TRN}           = 1
  
  IF newFeq{PROP:OLE} = FALSE AND newFeq{PROP:Object} = ''
    ! Creation failed, this is where you would attempt to register the DLL
    Stop('Unable to create the object')
    RETURN newFeq
  END

  IF newFeq{PROP:OLE} = FALSE AND newFeq{PROP:Object} <> ''
    Stop('PROP:Object=' & newFeq{PROP:Object} & ' but PROP:OLE=FALSE!')
    RETURN newFeq
  END
  
  Stop('Object and OLE container ready to go!')
  RETURN newFeq

  