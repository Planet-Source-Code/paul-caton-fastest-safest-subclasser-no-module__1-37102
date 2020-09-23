
;#########################################################################################################################
;#
;# This code represents the model used for cSuperClass message filtered subclassing.
;# We assemble this code merely to discover the opcodes to use in cSuperClass.cls
;# 
;# Paul_Caton@hotmail.com
;# 13th June 2002
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
    LOCAL   lHandled    :DWORD

    jmp     TestMsgNo               ;# Jump over the *constant* code to the run-time generated message number testing code

BeforePrevWndProc:                  ;# Jump here if we're handling this message before the previous WndProc
    mov     lReturn,0
    lea     eax,lReturn
    push    eax
    mov     lHandled,0
    lea     eax,lHandled
    push    eax
    mov     eax,88888888h           ;# Patched with ObjPtr(Owner) at run-time
    mov     ecx,eax
    mov     ecx,dword ptr [ecx]
    push    eax
    call    dword ptr [ecx+20h]     ;# Call Sub iSuperClass_Before

    cmp     lHandled,0              ;# Check to see if the user doesn't want the previous WndProc to receive this message
    jnz     Bail_1

    push    lParam                  ;# Call previous WndProc handler
    push    wParam
    push    uMsg
    push    hWin
    call    PrevWndProc             ;# PrevWndProc will be patched with the EIP relative offset to the real PrevWndProc at run-time
    ret
    
AfterPrevWndProc:                   ;# Jump here if we're handling this message number after the previous WndProc
    call    PrevWndProc             ;# PrevWndProc will be patched with the EIP relative offset to the real PrevWndProc at run-time
    mov     lReturn,eax

    push    lParam                  ;# Call our handler
    push    wParam
    push    uMsg
    push    hWin
    lea     eax,lReturn
    push    eax
    mov     eax,88888888h           ;# Patched with ObjPtr(Owner) at run-time
    mov     ecx,eax
    mov     ecx,dword ptr [ecx]
    push    eax
    call    dword ptr [ecx+1Ch]     ;# Call Sub iSuperClass_After
Bail_1:
    mov     eax,lReturn
    ret

TestMsgNo:
    mov     eax,uMsg

    push    lParam                  ;# No matter which way we branch these parameters are going to be stacked...
    push    wParam
    push    eax
    push    hWin
    
    cmp     eax,0BEF00000h          ;# The compare test and jump are dynamically added at run-time
    je      BeforePrevWndProc    
    
    cmp     eax,0AF000000h          ;# The compare test and jump are dynamically added at run-time
    je      AfterPrevWndProc     

    ;#
    ;# And so on for each message number added by the user
    ;#
    ;# Unspecified messages drop thru to here...
    call    PrevWndProc             ;# We're not interested in this message number, pass to the pre-existing window proc
    ret

PrevWndProc:
WndProc endp
end start