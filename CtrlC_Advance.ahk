
; CtrlC Advance - Una herramienta desarrollada para copiar varios elementos de texto
; y poder ser pegada en el orden con el que fueron agregadas a memoria. Para poder pegar
; el texto almacenado en memoria, presionar <ctrl+1> , <ctrl+2>, ... ,<ctrl+9> segun
; requiera.

MsgBox CtrlC Advance: Seleccione y pulse <ctrl+q> para agregar texto a la memoria y <ctrl+1>,...,<ctrl+9> para pegar el texto. [Deiner Zapata]

MyArray := []
Index := 0
MaxIndex = 9


;Esc::
^q::
ClipSaved := ClipboardAll  ; save the entire clipboard to the variable ClipSaved
clipboard := "" ; empty clipboard
Send, ^c        ; copy the selected word
ClipWait 1      ; wait for the clipboard to contain data
If !ErrorLevel  ; If NOT ErrorLevel clipwait found data on the clipboard
{
    Index++     ; checks the number in the variable "Index" and increases it by 1, each time you press esc.
    if (Index = MaxIndex+1) ; when the specific amount of words is exceeded
    {
    	msgbox Copia no almacenada, memoria llena.
    }
    MyArray.Insert(Index, clipboard)
	TrayTip, , Ctrl+%Index%, 3,1
}
Sleep, 300
clipboard := ClipSaved  ; restore original clipboard
return

^0::
Sleep, 300
clipboard := ClipSaved  ; restore original clipboard
MyArray := []
Index := 0
msgbox H23: Mem del clipboard limpiada.

return

^1:: SendInput, % MyArray[1]
^2:: SendInput, % MyArray[2]
^3:: SendInput, % MyArray[3]
^4:: SendInput, % MyArray[4]
^5:: SendInput, % MyArray[5]
^6:: SendInput, % MyArray[6]
^7:: SendInput, % MyArray[7]
^8:: SendInput, % MyArray[8]
^9:: SendInput, % MyArray[9]
