/*

NOTE: The code in its current state is a VERY rough MVP. Additionally, my attribution
may be incomplete.

Includes:
  VA.ahk - https://autohotkey.com/board/topic/21984-vista-audio-control-functions/

Borrowed Code:
  https://autohotkey.com/board/topic/54920-midi-inputoutput-combined-with-system-exclusive/
  https://github.com/micahstubbs/midi4ahk

*/

;;"#defines"
DeviceID := 0
CALLBACK_WINDOW := 0x10000


#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
#Persistent


Gui, +LastFound
hWnd := WinExist()


;MsgBox, hWnd = %hWnd%`nPress OK to open winmm.dll library

OpenCloseMidiAPI()
OnExit, Sub_Exit


;MsgBox, winmm.dll loaded.`nPress OK to open midi device`nDevice ID = %DeviceID%`nhWnd = %hWnd%`ndwFlags = CALLBACK_WINDOW

hMidiIn =
VarSetCapacity(hMidiIn, 4, 0)

result := DllCall("winmm.dll\midiInOpen", UInt,&hMidiIn, UInt,DeviceID, UInt,hWnd, UInt,0, UInt,CALLBACK_WINDOW, "UInt")

If result
{
    MsgBox, error, midiInOpen returned %result%`n
    GoSub, sub_exit
}

hMidiIn := NumGet(hMidiIn) ; because midiInOpen writes the value in 32 bit binary number, AHK stores it as a string

h_midiout := midiOutOpen(1)

Loop, 8 {
    midi_send(A_Index + 31, 0)
    midi_send(A_Index + 47, 0)
    midi_send(A_Index + 63, 0)
}
media_keys := [43, 44, 42, 41, 45]
for k, v in media_keys
    midi_send(v, 0)

midi_send(71, 127)

applications := []

IfExist, nano.ini
{
    loop, 8 {
        index := A_Index - 1
        IniRead, app, nano.ini, applications, %index%, ""
        applications[index] := app
        if (app != "")
            midi_send(index + 32, 127)
    }
}
else
{
    applications := ["","","","","","","",""]
}

meter_mode := false

;MsgBox, Midi input device opened successfully`nhMidiIn = %hMidiIn%`n`nPress OK to start the midi device

result := DllCall("winmm.dll\midiInStart", UInt,hMidiIn)
If result
{
    MsgBox, error, midiInStart returned %result%`n
    GoSub, sub_exit
}


;   #define MM_MIM_OPEN         0x3C1           /* MIDI input */
;   #define MM_MIM_CLOSE        0x3C2
;   #define MM_MIM_DATA         0x3C3
;   #define MM_MIM_LONGDATA     0x3C4
;   #define MM_MIM_ERROR        0x3C5
;   #define MM_MIM_LONGERROR    0x3C6

OnMessage(0x3C1, "midiInHandler")
OnMessage(0x3C2, "midiInHandler")
OnMessage(0x3C3, "midiInHandler")
OnMessage(0x3C4, "midiInHandler")
OnMessage(0x3C5, "midiInHandler")
OnMessage(0x3C6, "midiInHandler")


; AUDIO METER STUFF
MeterLength = 8

audioMeter := VA_GetAudioMeter()

VA_IAudioMeterInformation_GetMeteringChannelCount(audioMeter, channelCount)

; "The peak value for each channel is recorded over one device
;  period and made available during the subsequent device period."
VA_GetDevicePeriod("capture", devicePeriod)

do_reload := False

Loop
{
    ; Get the peak value across all channels.
    VA_IAudioMeterInformation_GetPeakValue(audioMeter, peakValue)
    meter := peakValue ;MakeMeter(peakValue, MeterLength)

    ; Get the peak values of all channels.
    ; VarSetCapacity(peakValues, channelCount*4)
    ; VA_IAudioMeterInformation_GetChannelsPeakValues(audioMeter, channelCount, &peakValues)
    ; Loop %channelCount%
    ;     meter .= "`n" MakeMeter(NumGet(peakValues, A_Index*4-4, "float"), MeterLength)

    midi_meter(peakValue)
    Sleep, %devicePeriod%
    if do_reload
        break
}

return

midi_meter(fraction) {
    global meter_mode
    if (meter_mode) {
        lights := [64, 65, 66, 67, 68, 69, 70, 71]
        size := 8
    }
    else {
        lights := [43, 44, 42, 41, 45]
        size := 5
    }



    light := Round(fraction*size)
;        midi_send(lights[A_Index-1], 127)
    for i, v in lights {
        if (i < light)
            midi_send(v, 127)
        Else
            midi_send(v, 0)
    }
    ;tooltip, %light%
}


sub_exit:

If (hMidiIn)
    DllCall("winmm.dll\midiInClose", UInt,hMidiIn)
OpenCloseMidiAPI()

ExitApp

;--------End of auto-execute section-----
;----------------------------------------


OpenCloseMidiAPI() {
   Static hModule
   If hModule
      DllCall("FreeLibrary", UInt,hModule), hModule := ""
   If (0 = hModule := DllCall("LoadLibrary",Str,"winmm.dll")) {
      MsgBox Cannot load library winmm.dll
      ExitApp
   }
}



midiInHandler(hInput, midiMsg, wMsg)
{
    statusbyte := midiMsg & 0xFF
    byte1 := (midiMsg >> 8) & 0xFF
    byte2 := (midiMsg >> 16) & 0xFF

;   MsgBox,
;   (
; Received a message: %wMsg%
; wParam = %hInput%
; lParam = %midiMsg%
;   statusbyte = %statusbyte%
;   byte1 = %byte1%
;   byte2 = %byte2%
;   )
    if (statusbyte == 176)
        midi_binding(byte1, byte2)
}

midi_binding(control, value)
{
    global meter_mode
    global applications
    spotify = ahk_class SpotifyMainWindow ;Set variable for Spotify Window Name
    ;MsgBox, %spotify%
    if (control == 41 && value = 127) {
        ;Send {Media_Play_Pause}
        ;ControlSend, ahk_parent, {Media_Play_Pause}, %spotify%
        Send {Media_Play_Pause}
    }
    else if (control == 45 && value = 127) {
        meter_mode := !meter_mode
        meter_controls := [43, 44, 42, 41, 45, 64, 65, 66, 67, 68, 69, 70, 71]
        for k, v in meter_controls
            midi_send(v, 0)
        if (!meter_mode)
            midi_send(71, 127)
    }
    else if (control == 42 && value = 127)
        Send {Media_Stop}
    else if (control == 43 && value = 127)
        Send {Media_Prev}
    else if (control == 44 && value = 127)
        Send {Media_Next}

    ; SET APPLICATION
    else if (control >= 32 && control <= 38 && value = 127) {
        appnum := control - 32
        if (applications[appnum] == "") {
            Winget,appname,ProcessName,A
            applications[appnum] := appname
            Tooltip, Fader %appnum% set to %appname%
            SetTimer, RemoveToolTip, 3000
            if (appname)
                midi_send(control, 127)
            WriteIni()
        }
        else
        {
            applications[appnum] := ""
            midi_send(control, 0)
            WriteIni()
        }
    }

    ; CHANGE MASTER VOLUME
    else if (control == 7)
        VA_SetMasterVolume(value/1.27)

    ; CHANGE APP VOLUMES
    else if (control >= 0 && control <= 7) {
        if (applications[control] == "")
            return
        app := applications[control]
        ;Tooltip, %control% = %app%
        SetTimer, RemoveToolTip, 1000
        Volume := GetVolumeObject(app)
        VA_ISimpleAudioVolume_SetMasterVolume(Volume, value/127)
    }

    ; MUTE MASTER VOLUME
    else if (control == 55 && value = 127) {
        Mute := VA_GetMasterMute()
        VA_SetMasterMute(!Mute)
        if (Mute)
            midi_send(control, 0)
        else
            midi_send(control, 127)
    }

    ; MUTE APP VOLUMES
    else if (control >= 48 && control <= 54 && value = 127) {
        appnum := control - 48
        if (applications[appnum] == "")
            return
        app := applications[appnum]
        Tooltip, %control% = %app%
        SetTimer, RemoveToolTip, 1000
        Volume := GetVolumeObject(app)
        VA_ISimpleAudioVolume_GetMute(Volume, Mute)
        VA_ISimpleAudioVolume_SetMute(Volume, !Mute)
        if (Mute)
            midi_send(control, 0)
        else
            midi_send(control, 127)
    }

    ; REVEAL APP
    else if (control >= 64 && control <= 70 && value = 127) {
        appnum := control - 64
        if (applications[appnum] == "")
            return
        app := applications[appnum]
        Tooltip, %appnum% = %app%
        SetTimer, RemoveToolTip, 1000
        WinActivate, ahk_exe %app%
    }

    ; REVEAL MIXER
    else if (control == 71 && value = 127) {
        If WinExist("ahk_exe SndVol.exe")
            WinActivate, ahk_exe SndVol.exe
        Else
            Run C:\Windows\System32\SndVol.exe
        Return
    }

    else if (control == 46 && value = 127) {
        global do_reload
        do_reload := true
        Reload
    }
;   else
;       ToolTip,
;       (
;Received a message:
;   control = %control%
;   value = %value%
;       )
}

RemoveToolTip:
SetTimer, RemoveToolTip, Off
ToolTip
return

^Esc::Reload
;Esc::GoSub, sub_exit


#include VA.ahk

; Get "System Sounds" (PID 0)
if !(Volume := GetVolumeObject("Spotify.exe"))
{
    MsgBox, There was a problem retrieving the application volume interface
    ExitApp
}
OnExit, ExitSub
return

ExitSub:
ObjRelease(Volume)
ExitApp
return


GetVolumeObject(Param)
{
    static IID_IASM2 := "{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}"
    , IID_IASC2 := "{bfb7ff88-7239-4fc9-8fa2-07c950be9c6d}"
    , IID_ISAV := "{87CE5498-68D6-44E5-9215-6DA47EF883D8}"

    ; Turn empty into integer
    if !Param
        Param := 0

    ; Get PID from process name
    if Param is not Integer
    {
        Process, Exist, %Param%
        Param := ErrorLevel
    }

    ; GetDefaultAudioEndpoint
    DAE := VA_GetDevice()

    ; activate the session manager
    VA_IMMDevice_Activate(DAE, IID_IASM2, 0, 0, IASM2)

    ; enumerate sessions for on this device
    VA_IAudioSessionManager2_GetSessionEnumerator(IASM2, IASE)
    VA_IAudioSessionEnumerator_GetCount(IASE, Count)

    ; search for an audio session with the required name
    Loop, % Count
    {
        ; Get the IAudioSessionControl object
        VA_IAudioSessionEnumerator_GetSession(IASE, A_Index-1, IASC)

        ; Query the IAudioSessionControl for an IAudioSessionControl2 object
        IASC2 := ComObjQuery(IASC, IID_IASC2)
        ObjRelease(IASC)

        ; Get the session's process ID
        VA_IAudioSessionControl2_GetProcessID(IASC2, SPID)

        ; If the process name is the one we are looking for
        if (SPID == Param)
        {
            ; Query for the ISimpleAudioVolume
            ISAV := ComObjQuery(IASC2, IID_ISAV)

            ObjRelease(IASC2)
            break
        }
        ObjRelease(IASC2)
    }
    ObjRelease(IASE)
    ObjRelease(IASM2)
    ObjRelease(DAE)
    return ISAV
}

;
; ISimpleAudioVolume : {87CE5498-68D6-44E5-9215-6DA47EF883D8}
;
VA_ISimpleAudioVolume_SetMasterVolume(this, ByRef fLevel, GuidEventContext="") {
    return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "float", fLevel, "ptr", VA_GUID(GuidEventContext))
}
VA_ISimpleAudioVolume_GetMasterVolume(this, ByRef fLevel) {
    return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "float*", fLevel)
}
VA_ISimpleAudioVolume_SetMute(this, ByRef Muted, GuidEventContext="") {
    return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "int", Muted, "ptr", VA_GUID(GuidEventContext))
}
VA_ISimpleAudioVolume_GetMute(this, ByRef Muted) {
    return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "int*", Muted)
}


midiOutOpen(uDeviceID = 0) { ; Open midi port for sending individual midi messages --> handle
strh_midiout = 0000

result := DllCall("winmm.dll\midiOutOpen", UInt,&strh_midiout, UInt,uDeviceID, UInt,0, UInt,0, UInt,0, UInt)
If (result or ErrorLevel) {
MsgBox There was an Error opening the midi port.`nError code %result%`nErrorLevel = %ErrorLevel%
Return -1
}
Return UInt@(&strh_midiout)
}

midiOutShortMsg(h_midiout, MidiStatus, Param1, Param2) { ;Channel,
;Tooltip, midiOutShortMsg %h_midiout% %MidiStatus% %Param1% %Param2%
;h_midiout: handle to midi output device returned by midiOutOpen
;EventType, Channel combined -> MidiStatus byte: http://www.harmony-central.com/MIDI/Doc/table1.html
;Param3 should be 0 for PChange, ChanAT, or Wheel
;Wheel events: entire Wheel value in Param2 - the function splits it into two bytes
/*
If (EventType = "NoteOn" OR EventType = "N1")
MidiStatus := 143 + Channel
Else If (EventType = "NoteOff" OR EventType = "N0")
MidiStatus := 127 + Channel
Else If (EventType = "CC")
MidiStatus := 175 + Channel
Else If (EventType = "PolyAT" OR EventType = "PA")
MidiStatus := 159 + Channel
Else If (EventType = "ChanAT" OR EventType = "AT")
MidiStatus := 207 + Channel
Else If (EventType = "PChange" OR EventType = "PC")
MidiStatus := 191 + Channel
Else If (EventType = "Wheel" OR EventType = "W") {
MidiStatus := 223 + Channel
Param2 := Param1 >> 8 ; MSB of wheel value
Param1 := Param1 & 0x00FF ; strip MSB
}
*/
result := DllCall("winmm.dll\midiOutShortMsg", UInt,h_midiout, UInt, MidiStatus|(Param1<<8)|(Param2<<16), UInt)
If (result or ErrorLevel) {
MsgBox There was an Error Sending the midi event: (%result%`, %ErrorLevel%)
Return -1
}
}

midi_send(control, value) {
    global h_midiout
    midiOutShortMsg(h_midiout, 176, control, value)
}



midiOutClose(h_midiout) { ; Close MidiOutput
Loop 9 {
result := DllCall("winmm.dll\midiOutClose", UInt,h_midiout)
If !(result or ErrorLevel)
Return
Sleep 250
}
MsgBox Error in closing the midi output port. There may still be midi events being Processed.
Return -1
}

;UTILITY FUNCTIONS
MidiOutGetNumDevs() { ; Get number of midi output devices on system, first device has an ID of 0
Return DllCall("winmm.dll\midiOutGetNumDevs")
}

MidiOutNameGet(uDeviceID = 0) { ; Get name of a midiOut device for a given ID

;MIDIOUTCAPS struct
; WORD wMid;
; WORD wPid;
; MMVERSION vDriverVersion;
; CHAR szPname[MAXPNAMELEN];
; WORD wTechnology;
; WORD wVoices;
; WORD wNotes;
; WORD wChannelMask;
; DWORD dwSupport;

VarSetCapacity(MidiOutCaps, 50, 0) ; allows for szPname to be 32 bytes
OffsettoPortName := 8, PortNameSize := 32
result := DllCall("winmm.dll\midiOutGetDevCapsA", UInt,uDeviceID, UInt,&MidiOutCaps, UInt,50, UInt)

If (result OR ErrorLevel) {
MsgBox Error %result% (ErrorLevel = %ErrorLevel%) in retrieving the name of midi output %uDeviceID%
Return -1
}

VarSetCapacity(PortName, PortNameSize)
DllCall("RtlMoveMemory", Str,PortName, Uint,&MidiOutCaps+OffsettoPortName, Uint,PortNameSize)
Return PortName
}

MidiOutsEnumerate() { ; Returns number of midi output devices, creates global array MidiOutPortName with their names
local NumPorts, PortID
MidiOutPortName =
NumPorts := MidiOutGetNumDevs()

Loop %NumPorts% {
PortID := A_Index -1
MidiOutPortName%PortID% := MidiOutNameGet(PortID)
}
Return NumPorts
}

UInt@(ptr) {
Return *ptr | *(ptr+1) << 8 | *(ptr+2) << 16 | *(ptr+3) << 24
}

PokeInt(p_value, p_address) { ; Windows 2000 and later
DllCall("ntdll\RtlFillMemoryUlong", UInt,p_address, UInt,4, UInt,p_value)
}

WriteIni()
{
    global applications

    IfNotExist, nano.ini ; if no ini
    FileAppend,, nano.ini ; make one with the following entries.
    for i, app in applications {
        IniWrite, %app%,nano.ini, applications, %i%
    }
}
