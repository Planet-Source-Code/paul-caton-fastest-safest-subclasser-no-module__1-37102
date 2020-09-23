
;#########################################################################################################################
;#
;# This code represents the model used for cSuperClass ALL message subclassing.
;# We assemble this code merely to discover the opcodes to use in cSuperClass.cls
;# 
;# Paul_Caton@hotmail.com
;# 29th June 2002
;#
;# P.S. I haven't assembled since the Atari ST... That's probably self-evident.
;#

.486                                ;# Create 32 bit code
.model flat, stdcall                ;# 32 bit memory model
option casemap :none                ;# Case sensitive
include WndProc.inc                 ;# Macros 'n stuff

.code

start:

WndProc proc    hWin    :DWORD,
                uMsg    :DWORD,
                wParam  :DWORD,
                lParam  :DWORD

    LOCAL   lReturn     :DWORD

    push    lParam                  
    push    wParam
    push    uMsg
    push    hWin
    call    PrevWndProc             ;# PrevWndProc will be patched with the EIP relative offset to the real PrevWndProc at run-time
    mov     lReturn,eax

    push    lParam                  ;# Call our handler
    push    wParam
    push    uMsg
    push    hWin
    lea     eax,lReturn
    push    eax
    mov     eax,8888888h            ;# Patched with ObjPtr(Owner) at run-time
    mov     ecx,eax
    mov     ecx,dword ptr [ecx]
    push    eax
    call    dword ptr [ecx+1Ch]     ;# Call Sub iSuperClass_After
    mov     eax,lReturn
    ret

PrevWndProc:
WndProc endp
end start