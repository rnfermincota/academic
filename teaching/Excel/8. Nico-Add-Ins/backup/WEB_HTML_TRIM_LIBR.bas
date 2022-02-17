Attribute VB_Name = "WEB_HTML_TRIM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function TRIM_HTML_TEXT_STRING_FUNC(ByVal HTML_DATA_STR As String, _
Optional ByVal IMAGE_TEXT_STR As String = "") As String
  
'Adapted from http://www.vbarchiv.de/

Dim i As Integer
Dim j As Integer

Dim CHR_STR As String
Dim HTML_STR As String
Dim HREF_STR As String

Dim PARAM_STR As String
Dim IMG_ALT_STR As String
Dim IMG_FILE_STR As String

Dim BLOCK_FLAG As Boolean

On Error Resume Next

' Pr�fen, ob <BODY>-Tag vorhanden
' wenn ja befindet sich der relevante HTML-Teil
' zwischen <BODY> und </BODY>
If InStr(LCase$(HTML_DATA_STR), "<body") > 0 Then
  HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(LCase$(HTML_DATA_STR), "<body"))
  HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(HTML_DATA_STR, ">") + 1)
  If InStr(LCase$(HTML_DATA_STR), "</body>") > 0 Then _
    HTML_DATA_STR = Left$(HTML_DATA_STR, InStr(LCase$(HTML_DATA_STR), _
    "</body>") - 1)
End If

' HTML-Text zeichenweise "scannen"
Do While Len(HTML_DATA_STR)
  CHR_STR = Left$(HTML_DATA_STR, 1)
  HTML_DATA_STR = Mid$(HTML_DATA_STR, 2)
  Select Case CHR_STR
    Case " "
      HTML_DATA_STR = LTrim$(HTML_DATA_STR)
    Case vbCr, vbLf
      If Right$(HTML_STR, 1) <> " " And _
        Right$(HTML_STR, 2) <> vbCrLf Then _
        CHR_STR = " " Else CHR_STR = ""
      If Left$(HTML_DATA_STR, 1) = vbLf Then _
        HTML_DATA_STR = Mid$(HTML_DATA_STR, 2)
      HTML_DATA_STR = LTrim$(HTML_DATA_STR)

    ' HTML-Steuerzeichen
    Case "<"
      CHR_STR = ""
      If InStr(HTML_DATA_STR, ">") > 0 Then
        CHR_STR = Left$(HTML_DATA_STR, InStr(HTML_DATA_STR, ">") - 1)
        HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(HTML_DATA_STR, ">") + 1)

        If InStr(CHR_STR, " ") > 0 Then
          PARAM_STR = REPLACE_CHR_FUNC(Trim$(Mid$(CHR_STR, _
            InStr(CHR_STR, " ") + 1)), vbCrLf, "")
          CHR_STR = Left$(CHR_STR, InStr(CHR_STR, " ") - 1)
        Else
          PARAM_STR = ""
        End If
        Dim ParamColl As Collection

        Select Case LCase$(CHR_STR)
          Case "p"
            If Right$(HTML_STR, 4) <> vbCrLf & vbCrLf Then _
              CHR_STR = vbCrLf & vbCrLf Else CHR_STR = ""
          Case "/div"
            If Right$(HTML_STR, 2) <> vbCrLf Then _
              CHR_STR = vbCrLf Else CHR_STR = ""
          Case "br"
            CHR_STR = vbCrLf
          Case "ul", "/ul", "ol", "/ol"
            CHR_STR = vbCrLf
          Case "li"
            CHR_STR = vbCrLf & "   - "
          Case "BLOCK_FLAG"
            BLOCK_FLAG = True
            CHR_STR = ""
          Case "/BLOCK_FLAG"
            BLOCK_FLAG = False
            CHR_STR = ""
          Case "img"
            If Len(IMAGE_TEXT_STR) > 0 Then
              Set ParamColl = STRIP_CHR_QUOTES_OBJ_FUNC(PARAM_STR, " ")
              IMG_ALT_STR = ""
              IMG_FILE_STR = ""
              For j = 1 To ParamColl.COUNT
                If LCase$(Left$(ParamColl(j), 4)) = "src=" Then
                  IMG_FILE_STR = STRIP_QUOTES_FUNC(Mid$(ParamColl(j), 5))
                ElseIf LCase$(Left$(ParamColl(j), 4)) = "alt=" Then
                  IMG_ALT_STR = STRIP_QUOTES_FUNC(Mid$(ParamColl(j), 5))
                End If
                If Len(IMG_ALT_STR) > 0 And Len(IMG_FILE_STR) > 0 Then Exit For
              Next
              If Len(Trim$(IMG_ALT_STR)) = 0 Then _
                IMG_ALT_STR = "Image " & IMG_FILE_STR
              IMAGE_TEXT_STR = REPLACE_CHR_FUNC(IMAGE_TEXT_STR, "%IMG_ALT_STR", IMG_ALT_STR)
              IMAGE_TEXT_STR = REPLACE_CHR_FUNC(IMAGE_TEXT_STR, "%imgsrc", IMG_FILE_STR)

              CHR_STR = " " & IMAGE_TEXT_STR & " "
            Else
              CHR_STR = ""
            End If
          Case "a"
            Set ParamColl = STRIP_CHR_QUOTES_OBJ_FUNC(PARAM_STR, " ")
            HREF_STR = ""
            For j = 1 To ParamColl.COUNT
              If LCase$(Left$(ParamColl(j), 5)) = "href=" Then
                HREF_STR = STRIP_QUOTES_FUNC(Mid$(ParamColl(j), 6))
                Exit For
              End If
            Next
            CHR_STR = ""
          Case "/a"
            CHR_STR = ""
            If Len(Trim$(HREF_STR)) > 0 Then
              If LCase$(Left$(HREF_STR, 7)) = "mailto:" Then _
                HREF_STR = Mid$(HREF_STR, 8)
              If LCase$(Right$(Trim$(HTML_STR), _
                Len(HREF_STR))) <> LCase$(HREF_STR) Then
                If Not (LCase$(Left$(HREF_STR, 7)) = "http://" And _
                 LCase$(Right$(Trim$(HTML_STR), _
                 Len(HREF_STR) - 7)) = LCase$(Mid$(HREF_STR, 8))) Then
                  CHR_STR = " [" & HREF_STR & "]"
                  HREF_STR = ""
                End If
              End If
            End If
          Case "hr"
            CHR_STR = vbCrLf
            For i = 1 To 70
              CHR_STR = CHR_STR & "-"
            Next
            CHR_STR = CHR_STR & vbCrLf
          Case "sigboundary"
            CHR_STR = vbCrLf & "-- " & vbCrLf
          Case "script"
            CHR_STR = ""
            If InStr(LCase$(HTML_DATA_STR), "</script>") > 0 Then _
              HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(LCase$(HTML_DATA_STR), _
                "</script>"))
          Case "pre"
            CHR_STR = vbCrLf
            If InStr(LCase$(HTML_DATA_STR), "</pre>") > 0 Then
              CHR_STR = CHR_STR & Left$(HTML_DATA_STR, _
                InStr(LCase$(HTML_DATA_STR), "</pre>") - 1)
              HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(LCase$(HTML_DATA_STR), _
                "</pre>"))
            End If
            CHR_STR = CHR_STR & vbCrLf
          Case Else
            CHR_STR = ""
        End Select
      End If
    Case "&"
      If InStr(HTML_DATA_STR, ";") > 0 And (InStr(HTML_DATA_STR, ";") < _
        InStr(HTML_DATA_STR, " ") Or InStr(HTML_DATA_STR, " ") = 0) Then
        CHR_STR = Left$(HTML_DATA_STR, InStr(HTML_DATA_STR, ";") - 1)
        HTML_DATA_STR = Mid$(HTML_DATA_STR, InStr(HTML_DATA_STR, ";") + 1)

        Select Case CHR_STR
          Case "amp"
            CHR_STR = "&"
          Case "quot"
            CHR_STR = """"
          Case "lt"
            CHR_STR = "<"
          Case "gt"
            CHR_STR = ">"
          Case "nbsp"
            CHR_STR = " "
          Case "Auml"
            CHR_STR = "�"
          Case "auml"
            CHR_STR = "�"
          Case "iexcl"
            CHR_STR = "�"
          Case "cent"
            CHR_STR = "�"
          Case "pound"
            CHR_STR = "�"
          Case "curren"
            CHR_STR = "�"
          Case "yen"
            CHR_STR = "�"
          Case "brvbar"
            CHR_STR = "|"
          Case "sect"
            CHR_STR = "�"
          Case "uml"
            CHR_STR = "�"
          Case "copy"
            CHR_STR = "�"
          Case "ordf"
            CHR_STR = "�"
          Case "laquo"
            CHR_STR = "�"
          Case "not"
            CHR_STR = "�"
          Case "reg"
            CHR_STR = "�"
          Case "macr"
            CHR_STR = "�"
          Case "deg"
            CHR_STR = "�"
          Case "plusm"
            CHR_STR = "�"
          Case "sup2"
            CHR_STR = "�"
          Case "sup3"
            CHR_STR = "�"
          Case "acute"
            CHR_STR = "�"
          Case "micro"
            CHR_STR = "�"
          Case "para"
            CHR_STR = "�"
          Case "middot"
            CHR_STR = "�"
          Case "cedil"
            CHR_STR = "�"
          Case "sup1"
            CHR_STR = "�"
          Case "ordm"
            CHR_STR = "�"
          Case "raquo"
            CHR_STR = "�"
          Case "frac14"
            CHR_STR = "�"
          Case "frac12"
            CHR_STR = "�"
          Case "frac34"
            CHR_STR = "�"
          Case "iquest"
            CHR_STR = "�"
          Case "Agrave"
            CHR_STR = "�"
          Case "Aacute"
            CHR_STR = "�"
          Case "Acirc"
            CHR_STR = "�"
          Case "Atilde"
            CHR_STR = "�"
          Case "Aring"
            CHR_STR = "�"
          Case "AElig"
            CHR_STR = "�"
          Case "Ccedil"
            CHR_STR = "�"
          Case "Egrave"
            CHR_STR = "�"
          Case "Eacute"
            CHR_STR = "�"
          Case "Ecirc"
            CHR_STR = "�"
          Case "Euml"
            CHR_STR = "�"
          Case "Igrave"
            CHR_STR = "�"
          Case "Iacute"
            CHR_STR = "�"
          Case "Icirc"
            CHR_STR = "�"
          Case "Iuml"
            CHR_STR = "�"
          Case "ETH"
            CHR_STR = "�"
          Case "Ntilde"
            CHR_STR = "�"
          Case "Ograve"
            CHR_STR = "�"
          Case "Oacute"
            CHR_STR = "�"
          Case "Ocirc"
            CHR_STR = "�"
          Case "Otilde"
            CHR_STR = "�"
          Case "Ouml"
            CHR_STR = "�"
          Case "times"
            CHR_STR = "�"
          Case "Oslash"
            CHR_STR = "�"
          Case "Ugrave"
            CHR_STR = "�"
          Case "Uacute"
            CHR_STR = "�"
          Case "Ucirc"
            CHR_STR = "�"
          Case "Uuml"
            CHR_STR = "�"
          Case "Yacute"
            CHR_STR = "�"
          Case "THORN"
            CHR_STR = "�"
          Case "szlig"
            CHR_STR = "�"
          Case "agrave"
            CHR_STR = "�"
          Case "aacute"
            CHR_STR = "�"
          Case "acirc"
            CHR_STR = "�"
          Case "atilde"
            CHR_STR = "�"
          Case "aring"
            CHR_STR = "�"
          Case "aelig"
            CHR_STR = "�"
          Case "ccedil"
            CHR_STR = "�"
          Case "egrave"
            CHR_STR = "�"
          Case "eacute"
            CHR_STR = "�"
          Case "ecirc"
            CHR_STR = "�"
          Case "euml"
            CHR_STR = "�"
          Case "igrave"
            CHR_STR = "�"
          Case "iacute"
            CHR_STR = "�"
          Case "icirc"
            CHR_STR = "�"
          Case "iuml"
            CHR_STR = "�"
          Case "eth"
            CHR_STR = "�"
          Case "ntilde"
            CHR_STR = "�"
          Case "ograve"
            CHR_STR = "�"
          Case "oacute"
            CHR_STR = "�"
          Case "ocirc"
            CHR_STR = "�"
          Case "otilde"
            CHR_STR = "�"
          Case "ouml"
            CHR_STR = "�"
          Case "divide"
            CHR_STR = "�"
          Case "oslash"
            CHR_STR = "�"
          Case "ugrave"
            CHR_STR = "�"
          Case "uacute"
            CHR_STR = "�"
          Case "ucirc"
            CHR_STR = "�"
          Case "uuml"
            CHR_STR = "�"
          Case "yacute"
            CHR_STR = "�"
          Case "thorn"
            CHR_STR = "�"
          Case "yuml"
            CHR_STR = "�"
          Case Else
            CHR_STR = "&" & CHR_STR & ";"
        End Select
      End If
  End Select
  If Right$(CHR_STR, 2) = vbCrLf And BLOCK_FLAG Then _
    CHR_STR = CHR_STR & "> "
  HTML_STR = HTML_STR & CHR_STR
Loop

HTML_STR = Trim$(HTML_STR)

Do While Left$(HTML_STR, 2) = vbCrLf
  HTML_STR = Trim$(Mid$(HTML_STR, 3))
Loop
Do While Right$(HTML_STR, 2) = vbCrLf
  HTML_STR = Trim$(Left$(HTML_STR, Len(HTML_STR) - 2))
Loop

TRIM_HTML_TEXT_STRING_FUNC = HTML_STR
End Function

Private Function REPLACE_CHR_FUNC(ByVal DATA_STR As String, _
ByVal REPLACE_STR As String, _
ByVal INSERT_STR As String) As String
 
Dim i As Long
Dim j As Long

i = 0
Do While InStr(i + 1, LCase$(DATA_STR), LCase$(REPLACE_STR)) <> 0
  j = InStr(i + 1, LCase$(DATA_STR), LCase$(REPLACE_STR))
  DATA_STR = Left$(DATA_STR, InStr(i + 1, LCase$(DATA_STR), _
    LCase$(REPLACE_STR)) - 1) & INSERT_STR & Mid$(DATA_STR, _
    InStr(i + 1, LCase$(DATA_STR), LCase$(REPLACE_STR)) + _
    Len(REPLACE_STR))
  i& = j + Len(INSERT_STR) - 1
Loop

REPLACE_CHR_FUNC = DATA_STR

Exit Function
ERROR_LABEL:
REPLACE_CHR_FUNC = Err.number
End Function

Private Function STRIP_CHR_QUOTES_OBJ_FUNC(ByVal DATA_STR As String, _
ByVal STRIP_CHR_STR As String) As Collection

Dim i As Long
Dim j As Long
Dim TEMP_OBJ As Collection

On Error GoTo ERROR_LABEL

Set TEMP_OBJ = New Collection
Do
  i = InStr(DATA_STR, STRIP_CHR_STR)
  j = InStr(DATA_STR, Chr$(34))

  If j > 0 And j < i Then
    j = InStr(j + 1, DATA_STR, Chr$(34))
    Do While j > i And j > 0 And i > 0
      i = InStr(i + 1, DATA_STR, STRIP_CHR_STR)
    Loop
  End If

  If i > 0 Then
    TEMP_OBJ.Add Left$(DATA_STR, i - 1)
    DATA_STR = Mid$(DATA_STR, i + 1)
  Else
    TEMP_OBJ.Add DATA_STR
    DATA_STR = ""
  End If
Loop While Len(DATA_STR) > 0

Set STRIP_CHR_QUOTES_OBJ_FUNC = TEMP_OBJ
Set TEMP_OBJ = Nothing
  
Exit Function
ERROR_LABEL:
Set STRIP_CHR_QUOTES_OBJ_FUNC = Nothing
End Function

Private Function STRIP_QUOTES_FUNC(ByVal TEXT_STR As String) As String

On Error GoTo ERROR_LABEL

If Left$(TEXT_STR, 1) = Chr$(34) Then TEXT_STR = Mid$(TEXT_STR, 2)
If Right$(TEXT_STR, 1) = Chr$(34) Then TEXT_STR = Left$(TEXT_STR, Len(TEXT_STR) - 1)

Exit Function
ERROR_LABEL:
STRIP_QUOTES_FUNC = Err.number
End Function


'HTML Characters
'  &#00;-&#08; Unused
'  &#09; \t Horizontal tab
'  &#10; \n Line feed
'  &#11;-&#12; Unused
'  &#13; \r Carriage Return
'  &#14;-&#31; Unused
'  &#32; \s Space
'  &#33; ! Exclamation mark
'  &#34; " Quotation mark
'  &#35; # Number sign
'  &#36; $ Dollar sign
'  &#37; % Percent sign
'  &#38;(&amp;) & Ampersand
'  &#39; ' Apostrophe
'  &#40; ( Left parenthesis
'  &#41; ) Right parenthesis
'  &#42; * Asterisk
'  &#43; + Plus sign
'  &#44; , Comma
'  &#45; - Hyphen
'  &#46; . Period (fullstop)
'  &#47; / Solidus (slash)
'  &#48;-&#57; 0-9 Digits 0-9
'  &#58; : Colon
'  &#59; ; Semi-colon
'  &#60;(&lt;) < Less than
'  &#61; = Equals sign
'  &#62;(&gt;) > Greater than
'  &#63; ? Question mark
'  &#64; @ Commercial at
'  &#65;-&#90; A-Z Letters A-Z
'  &#91; [ Left square bracket
'  &#92; \ Reverse solidus (backslash)
'  &#93; ] Right square bracket
'  &#94; ^ Caret
'  &#95; _ Horizontal bar (underscore)
'  &#96; ` Acute accent
'  &#97;-&#122; a-z Letters a-z
'  &#123; { Left curly brace
'  &#124; | Vertical bar
'  &#125; } Right curly brace
'  &#126; ~ Tilde
'  &#127;-&#159; Unused
'  &#160;(&nbsp;) \s(nb) Non-breaking Space
'  &#161; � Inverted exclamation
'  &#162; � Cent sign
'  &#163; � Pound sterling
'  &#164; � General currency sign
'  &#165; � Yen sign
'  &#166; � Broken vertical bar
'  &#167; � Section sign
'  &#168; � Umlaut (dieresis)
'  &#169;(&copy;) � Copyright
'  &#170; � Feminine ordinal
'  &#171; � Left angle quote, guillemotleft
'  &#172; � Not sign
'  &#173; � Soft hyphen
'  &#174; � Registered trademark
'  &#175; � Macron accent
'  &#176; � Degree sign
'  &#177; � Plus or minus
'  &#178; � Superscript two
'  &#179; � Superscript three
'  &#180; � Acute accent
'  &#181; � Micro sign
'  &#182; � Paragraph sign
'  &#183; � Middle dot
'  &#184; � Cedilla
'  &#185; � Superscript one
'  &#186; � Masculine ordinal
'  &#187; � Right angle quote, guillemotright
'  &#188;(&frac14;) � Fraction one-fourth
'  &#189;(&frac12;) � Fraction one-half
'  &#190;(&frac34;) � Fraction three-fourths
'  &#191; � Inverted question mark
'  &#192; � Capital A, grave accent
'  &#193; � Capital A, acute accent
'  &#194; � Capital A, circumflex accent
'  &#195; � Capital A, tilde
'  &#196; � Capital A, dieresis or umlaut mark
'  &#197; � Capital A, ring
'  &#198; � Capital AE dipthong (ligature)
'  &#199; � Capital C, cedilla
'  &#200; � Capital E, grave accent
'  &#201; � Capital E, acute accent
'  &#202; � Capital E, circumflex accent
'  &#203; � Capital E, dieresis or umlaut mark
'  &#204; � Capital I, grave accent
'  &#205; � Capital I, acute accent
'  &#206; � Capital I, circumflex accent
'  &#207; � Capital I, dieresis or umlaut mark
'  &#208; � Capital Eth, Icelandic
'  &#209; � Capital N, tilde
'  &#210; � Capital O, grave accent
'  &#211; � Capital O, acute accent
'  &#212; � Capital O, circumflex accent
'  &#213; � Capital O, tilde
'  &#214; � Capital O, dieresis or umlaut mark
'  &#215; � Multiply sign
'  &#216; � Capital O, slash
'  &#217; � Capital U, grave accent
'  &#218; � Capital U, acute accent
'  &#219; � Capital U, circumflex accent
'  &#220; � Capital U, dieresis or umlaut mark
'  &#221; � Capital Y, acute accent
'  &#222; � Capital THORN, Icelandic
'  &#223; � Small sharp s, German (sz ligature)
'  &#224; � Small a, grave accent
'  &#225; � Small a, acute accent
'  &#226; � Small a, circumflex accent
'  &#227; � Small a, tilde
'  &#228; � Small a, dieresis or umlaut mark
'  &#229; � Small a, ring
'  &#230; � Small ae dipthong (ligature)
'  &#231; � Small c, cedilla
'  &#232; � Small e, grave accent
'  &#233; � Small e, acute accent
'  &#234; � Small e, circumflex accent
'  &#235; � Small e, dieresis or umlaut mark
'  &#236; � Small i, grave accent
'  &#237; � Small i, acute accent
'  &#238; � Small i, circumflex accent
'  &#239; � Small i, dieresis or umlaut mark
'  &#240; � Small eth, Icelandic
'  &#241; � Small n, tilde
'  &#242; � Small o, grave accent
'  &#243; � Small o, acute accent
'  &#244; � Small o, circumflex accent
'  &#245; � Small o, tilde
'  &#246; � Small o, dieresis or umlaut mark
'  &#247; � Division sign
'  &#248; � Small o, slash
'  &#249; � Small u, grave accent
'  &#250; � Small u, acute accent
'  &#251; � Small u, circumflex accent
'  &#252; � Small u, dieresis or umlaut mark
'  &#253; � Small y, acute accent
'  &#254; � Small thorn, Icelandic
'  &#255; � Small y, dieresis or umlaut mark
