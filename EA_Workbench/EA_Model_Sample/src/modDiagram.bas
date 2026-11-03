Attribute VB_Name = "modDiagram"
' =====================================================================
' draw.io integration (phase 4). draw.io Desktop and the Workbench
' exchange data only through files in the model folder:
'   Library\ArchiMate4.xml     palette of tagged ArchiMate 4 shapes and connectors
'   Library\ModelLibrary.xml   every active entity as a ready-tagged shape
'   Diagrams\*.drawio          uncompressed diagrams; shapes carry ea_id / ea_type,
'                              connectors carry ea_rel_id / ea_reltype
' Generate writes or extends a page, Update refreshes labels and marks missing
' entities, Harvest reads diagrams back (preview, then commit) and writes the
' new IDs into the file. SYS_Views registers every page with a fingerprint.
' The layout and the library format mirror build/drawio.py.
' =====================================================================
Option Explicit

Private Const SH_FORM As String = "FORM_Diagram"
Private Const SH_REVIEW As String = "HARVEST_Review"
Private Const SH_VIEWS As String = "SYS_Views"
Private Const BAND_ORDER As String = "Motivation|Strategy|Business|Common|Application|Technology|Implementation & Migration"
Private Const SHAPE_W As Long = 160
Private Const SHAPE_H As Long = 70
Private Const GAP_X As Long = 40
Private Const GAP_Y As Long = 30
Private Const GRID_COLS As Long = 6
Private Const LEFT_X As Long = 40
Private Const TOP_Y As Long = 40
Private Const BAND_STYLE As String = "text;html=1;align=left;verticalAlign=middle;fontStyle=1;fontColor=#595959;ea_band=1;"
Private Const DEFAULT_ENTITY_STYLE As String = "rounded=1;whiteSpace=wrap;html=1;"
Private Const DEFAULT_REL_STYLE As String = "html=1;rounded=0;endArrow=open;endFill=0;"

Private mStyles As Object        ' StyleKey -> Array(FillColour, DrawioStyle)
Private mProblems As String

' ================================================================ paths
Private Function ModelFolder() As String
    ModelFolder = ThisWorkbook.Path & Application.PathSeparator
End Function

Private Function SubFolder(ByVal nm As String) As String
    Dim p As String, part As Variant, acc As String
    acc = ModelFolder()
    For Each part In Split(nm, "\")
        acc = acc & part & Application.PathSeparator
        If Dir(acc, vbDirectory) = "" Then MkDir acc
    Next part
    SubFolder = acc
End Function

Public Function DiagramsFolder() As String
    DiagramsFolder = SubFolder("Diagrams")
End Function

Private Function SafeFile(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, CStr(ch), "-")
    Next ch
    SafeFile = Trim$(s)
End Function

' File as stored in SYS_Views: the name when it is in Diagrams\, else the full path
Private Function ViewFile(ByVal path As String) As String
    If LCase$(Left$(path, Len(DiagramsFolder()))) = LCase$(DiagramsFolder()) Then
        ViewFile = Mid$(path, Len(DiagramsFolder()) + 1)
    Else
        ViewFile = path
    End If
End Function

Private Function FullPath(ByVal viewFileName As String) As String
    If InStr(viewFileName, ":") > 0 Or Left$(viewFileName, 2) = "\\" Then
        FullPath = viewFileName
    Else
        FullPath = DiagramsFolder() & viewFileName
    End If
End Function

Private Function FileExists(ByVal path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    On Error Resume Next
    FileExists = (Len(Dir(path)) > 0)
End Function

Private Function FileStamp(ByVal path As String) As String
    On Error Resume Next
    FileStamp = Format$(FileDateTime(path), "yyyymmddhhnnss")
End Function

' ================================================================ styles
Private Sub LoadStyles()
    Dim d As Variant, i As Long
    Set mStyles = modCore.NewDict()
    d = modCore.TableData(modCore.Tbl("MM_Styles", "tblStyles"))
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, 1)) <> "" Then mStyles(CStr(d(i, 1))) = Array(CStr(d(i, 3)), CStr(d(i, 4)))
        Next i
    End If
End Sub

' Default draw.io style for a type added with the Add Type wizard (modForms.EnsureStyle)
Public Function DefaultEntityStyle() As String
    DefaultEntityStyle = DEFAULT_ENTITY_STYLE
End Function

Public Function EntityStyle(ByVal typeCode As String) As String
    Dim key As String, st As String, fill As String
    If mStyles Is Nothing Then LoadStyles
    key = "Entity:" & typeCode
    If modMeta.gTypes.Exists(typeCode) Then
        If modMeta.gTypes(typeCode)("StyleKey") <> "" Then key = modMeta.gTypes(typeCode)("StyleKey")
    End If
    If mStyles.Exists(key) Then
        fill = mStyles(key)(0): st = mStyles(key)(1)
    End If
    If InStr(st, "=") = 0 Then st = DEFAULT_ENTITY_STYLE      ' placeholder such as custom:Code
    If Right$(st, 1) <> ";" Then st = st & ";"
    If InStr(st, "fillColor=") = 0 And fill <> "" Then st = st & "fillColor=#" & fill & ";"
    EntityStyle = st
End Function

Public Function RelStyle(ByVal relCode As String) As String
    Dim st As String
    If mStyles Is Nothing Then LoadStyles
    If mStyles.Exists("Rel:" & relCode) Then st = mStyles("Rel:" & relCode)(1)
    If InStr(st, "=") = 0 Then st = DEFAULT_REL_STYLE
    If Right$(st, 1) <> ";" Then st = st & ";"
    RelStyle = st
End Function

Private Function StyleValue(ByVal style As String, ByVal key As String) As String
    Dim part As Variant, p As Long
    For Each part In Split(style, ";")
        p = InStr(part, "=")
        If p > 0 Then
            If Trim$(Left$(part, p - 1)) = key Then StyleValue = Trim$(Mid$(part, p + 1))
        End If
    Next part
End Function

Private Function SetStyleKey(ByVal style As String, ByVal key As String, ByVal value As String) As String
    Dim parts As Variant, i As Long, p As Long, out As String, done As Boolean
    parts = Split(style, ";")
    For i = LBound(parts) To UBound(parts)
        If parts(i) <> "" Then
            p = InStr(parts(i), "=")
            If p > 0 Then
                If Trim$(Left$(parts(i), p - 1)) = key Then
                    If value <> "" Then out = out & key & "=" & value & ";"
                    done = True
                Else
                    out = out & parts(i) & ";"
                End If
            Else
                out = out & parts(i) & ";"
            End If
        End If
    Next i
    If Not done And value <> "" Then out = out & key & "=" & value & ";"
    SetStyleKey = out
End Function

' Arrow and dash keys with draw.io's defaults; identifies a relationship from a copied line style
Private Function Signature(ByVal style As String) As String
    Dim sa As String, sf As String, ea As String, ef As String, dsh As String, dp As String
    sa = StyleValue(style, "startArrow"): If sa = "" Then sa = "none"
    ea = StyleValue(style, "endArrow"): If ea = "" Then ea = "classic"
    sf = StyleValue(style, "startFill"): If sf = "" Then sf = "1"
    ef = StyleValue(style, "endFill"): If ef = "" Then ef = "1"
    dsh = StyleValue(style, "dashed"): If dsh = "" Then dsh = "0"
    dp = StyleValue(style, "dashPattern")
    If sa = "none" Then sf = "-"
    If ea = "none" Then ef = "-"
    If dsh <> "1" Then dp = ""
    Signature = "startArrow=" & sa & ";startFill=" & sf & ";endArrow=" & ea & ";endFill=" & ef & _
                ";dashed=" & dsh & ";dashPattern=" & dp
End Function

Private Function RelFromStyle(ByVal style As String) As String
    Dim rc As Variant, sig As String, found As String, n As Long
    If mStyles Is Nothing Then LoadStyles
    sig = Signature(style)
    For Each rc In modMeta.gRelTypes.Keys
        If modMeta.gRelTypes(rc)("Status") = "Active" And mStyles.Exists("Rel:" & rc) Then
            If Signature(RelStyle(CStr(rc))) = sig Then found = CStr(rc): n = n + 1
        End If
    Next rc
    If n = 1 Then RelFromStyle = found
End Function

' ================================================================ compression (library entries)
' encodeURIComponent: UTF-8 bytes, percent-encoded except A-Z a-z 0-9 - _ . ! ~ * ' ( )
Public Function EncodeURIComponent(ByVal s As String) As String
    Dim i As Long, u As Long, lo As Long, cp As Long, out As String, ch As String
    Dim parts() As String, n As Long
    ReDim parts(1 To Len(s) + 1)
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        u = AscW(ch): If u < 0 Then u = u + 65536
        If (u >= 65 And u <= 90) Or (u >= 97 And u <= 122) Or (u >= 48 And u <= 57) Or _
           InStr("-_.!~*'()", ch) > 0 Then
            out = ch
        Else
            cp = u
            If u >= &HD800& And u <= &HDBFF& And i < Len(s) Then
                lo = AscW(Mid$(s, i + 1, 1)): If lo < 0 Then lo = lo + 65536
                If lo >= &HDC00& And lo <= &HDFFF& Then
                    cp = &H10000 + (u - &HD800&) * 1024& + (lo - &HDC00&)
                    i = i + 1
                End If
            End If
            If cp < &H80& Then
                out = Pct(cp)
            ElseIf cp < &H800& Then
                out = Pct(&HC0& Or (cp \ 64)) & Pct(&H80& Or (cp And 63))
            ElseIf cp < &H10000 Then
                out = Pct(&HE0& Or (cp \ 4096)) & Pct(&H80& Or ((cp \ 64) And 63)) & Pct(&H80& Or (cp And 63))
            Else
                out = Pct(&HF0& Or (cp \ 262144)) & Pct(&H80& Or ((cp \ 4096) And 63)) & _
                      Pct(&H80& Or ((cp \ 64) And 63)) & Pct(&H80& Or (cp And 63))
            End If
        End If
        n = n + 1
        parts(n) = out
        i = i + 1
    Loop
    If n = 0 Then Exit Function
    ReDim Preserve parts(1 To n)
    EncodeURIComponent = Join(parts, "")
End Function

Private Function Pct(ByVal b As Long) As String
    Pct = "%" & Right$("0" & Hex$(b), 2)
End Function

' base64(deflateRaw(encodeURIComponent(xml))) using stored deflate blocks (valid deflate, no compressor needed)
Public Function Compress(ByVal xml As String) As String
    Dim src() As Byte, out() As Byte, n As Long, i As Long, pos As Long, chunk As Long, j As Long
    Dim enc As String, el As Object, dom As Object
    enc = EncodeURIComponent(xml)
    n = Len(enc)
    If n = 0 Then
        ReDim out(0 To 4)
        out(0) = 1: out(1) = 0: out(2) = 0: out(3) = 255: out(4) = 255
    Else
        src = StrConv(enc, vbFromUnicode)
        ReDim out(0 To n + 5 * ((n - 1) \ 65535 + 1) - 1)
        i = 0: pos = 0
        Do While i < n
            chunk = n - i
            If chunk > 65535 Then chunk = 65535
            out(pos) = IIf(i + chunk >= n, 1, 0)
            out(pos + 1) = chunk And 255
            out(pos + 2) = (chunk \ 256) And 255
            out(pos + 3) = (65535 - chunk) And 255
            out(pos + 4) = (65535 - chunk) \ 256
            pos = pos + 5
            For j = 0 To chunk - 1
                out(pos + j) = src(i + j)
            Next j
            pos = pos + chunk
            i = i + chunk
        Loop
    End If
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set el = dom.createElement("b")
    el.DataType = "bin.base64"
    el.nodeTypedValue = out
    Compress = Replace(Replace(el.Text, vbLf, ""), vbCr, "")
End Function

' ================================================================ xml helpers
Private Function NewDom() As Object
    Dim dom As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.validateOnParse = False
    dom.resolveExternals = False
    dom.preserveWhiteSpace = False
    dom.setProperty "SelectionLanguage", "XPath"
    Set NewDom = dom
End Function

Private Function XmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    XmlEsc = s
End Function

Private Function JsonStr(ByVal s As String) As String
    Dim i As Long, ch As String, u As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        u = AscW(ch): If u < 0 Then u = u + 65536
        Select Case True
            Case ch = """": out = out & "\"""
            Case ch = "\": out = out & "\\"
            Case u < 32: out = out & "\u" & Right$("000" & Hex$(u), 4)
            Case Else: out = out & ch
        End Select
    Next i
    JsonStr = """" & out & """"
End Function

Private Function Attr(node As Object, ByVal nm As String) As String
    Dim v As Variant
    If node Is Nothing Then Exit Function
    v = node.getAttribute(nm)
    If Not IsNull(v) Then Attr = CStr(v)
End Function

' The mxCell of a draw.io cell node (<object>/<UserObject> wrap it; a plain <mxCell> is its own cell)
Private Function InnerCell(node As Object) As Object
    If node.nodeName = "mxCell" Then
        Set InnerCell = node
    Else
        Set InnerCell = node.selectSingleNode("mxCell")
    End If
End Function

Private Function IsTaggable(node As Object) As Boolean
    IsTaggable = (node.nodeName = "object" Or node.nodeName = "UserObject")
End Function

Private Function CellLabel(node As Object) As String
    If IsTaggable(node) Then CellLabel = Attr(node, "label") Else CellLabel = Attr(node, "value")
End Function

' Label text without HTML formatting
Private Function CleanLabel(ByVal s As String) As String
    Dim out As String, i As Long, ch As String, inTag As Boolean, tagText As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If inTag Then
            If ch = ">" Then
                inTag = False
                If LCase$(Left$(tagText, 2)) = "br" Or LCase$(Left$(tagText, 3)) = "div" Or _
                   LCase$(Left$(tagText, 4)) = "/div" Or LCase$(Left$(tagText, 1)) = "p" Then out = out & " "
            Else
                tagText = tagText & ch
            End If
        ElseIf ch = "<" Then
            inTag = True: tagText = ""
        Else
            out = out & ch
        End If
    Next i
    out = Replace(out, "&nbsp;", " ")
    out = Replace(out, "&lt;", "<")
    out = Replace(out, "&gt;", ">")
    out = Replace(out, "&quot;", """")
    out = Replace(out, "&#39;", "'")
    out = Replace(out, "&amp;", "&")
    out = Replace(out, vbCr, " "): out = Replace(out, vbLf, " "): out = Replace(out, ChrW$(160), " ")
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    CleanLabel = Trim$(out)
End Function

' HTML text for a label in an html=1 cell
Private Function HtmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    HtmlEsc = Replace(s, ">", "&gt;")
End Function

' A snapshot of a node's element children, safe to iterate while the tree changes
Private Function Children(root As Object) As Collection
    Dim c As New Collection, node As Object
    For Each node In root.childNodes
        If node.nodeType = 1 Then c.Add node
    Next node
    Set Children = c
End Function

' The .drawio files in Diagrams (read in full before any other Dir call)
Private Function DiagramFiles() As Collection
    Dim c As New Collection, folder As String, f As String
    folder = DiagramsFolder()
    f = Dir(folder & "*.drawio")
    Do While f <> ""
        c.Add f
        f = Dir()
    Loop
    Set DiagramFiles = c
End Function

Private Function NewCellID() As String
    NewCellID = "ea" & LCase$(Left$(Replace(Replace(Replace(modCore.NewGUID(), "-", ""), "{", ""), "}", ""), 16))
End Function

' Loads a .drawio file; returns Nothing with msg set when it cannot be read
Private Function LoadDrawio(ByVal path As String, ByRef msg As String) As Object
    Dim dom As Object
    Set dom = NewDom()
    If Not dom.Load(path) Then
        msg = "Not a readable draw.io file (" & Trim$(dom.parseError.reason) & ")"
        Exit Function
    End If
    If dom.documentElement Is Nothing Then msg = "Empty file": Exit Function
    If dom.documentElement.nodeName <> "mxfile" Then
        msg = "Not a draw.io diagram file (no mxfile element)"
        Exit Function
    End If
    Set LoadDrawio = dom
End Function

Private Function NewMxfile() As Object
    Dim dom As Object, root As Object
    Set dom = NewDom()
    dom.appendChild dom.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    Set root = dom.createElement("mxfile")
    root.setAttribute "host", "EA Workbench"
    dom.appendChild root
    Set NewMxfile = dom
End Function

Private Function FindPage(dom As Object, ByVal pageName As String) As Object
    Dim p As Object
    For Each p In dom.documentElement.selectNodes("diagram")
        If StrComp(Attr(p, "name"), pageName, vbTextCompare) = 0 Then Set FindPage = p: Exit Function
    Next p
End Function

Private Function AddPage(dom As Object, ByVal pageName As String) As Object
    Dim pg As Object, m As Object, root As Object, c As Object, kv As Variant
    Set pg = dom.createElement("diagram")
    pg.setAttribute "id", NewCellID()
    pg.setAttribute "name", pageName
    Set m = dom.createElement("mxGraphModel")
    For Each kv In Array("dx=1400", "dy=900", "grid=1", "gridSize=10", "guides=1", "tooltips=1", "connect=1", _
                         "arrows=1", "fold=1", "page=1", "pageScale=1", "pageWidth=1654", "pageHeight=1169", _
                         "math=0", "shadow=0")
        m.setAttribute Split(kv, "=")(0), Split(kv, "=")(1)
    Next kv
    Set root = dom.createElement("root")
    Set c = dom.createElement("mxCell"): c.setAttribute "id", "0": root.appendChild c
    Set c = dom.createElement("mxCell"): c.setAttribute "id", "1": c.setAttribute "parent", "0": root.appendChild c
    m.appendChild root
    pg.appendChild m
    dom.documentElement.appendChild pg
    Set AddPage = pg
End Function

' The root element of a page; Nothing when the page is compressed
Private Function PageRoot(pg As Object) As Object
    Dim m As Object
    Set m = pg.selectSingleNode("mxGraphModel")
    If m Is Nothing Then Exit Function
    Set PageRoot = m.selectSingleNode("root")
End Function

Private Function CompressedMsg(ByVal pageName As String) As String
    CompressedMsg = "Page """ & pageName & """ is saved compressed. In draw.io untick File > Properties > Compressed, " & _
                    "save, then try again."
End Function

Private Sub SaveDom(dom As Object, ByVal path As String)
    dom.Save path
End Sub

' ================================================================ repository data
' EntityID -> Array(TypeCode, Name, DisplayID, Status, RowHash)
Private Function MasterInfo() As Object
    Dim d As Variant, i As Long, cols As Object, lo As ListObject, b As Long, info As Object
    Set info = modCore.NewDict()
    Set lo = modData.MasterLO(): Set cols = modData.MasterCols()
    d = modCore.TableData(lo)
    b = lo.Range.Column - 1
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, cols("EntityID") - b)) <> "" Then
                info(CStr(d(i, cols("EntityID") - b))) = Array(CStr(d(i, cols("TypeCode") - b)), _
                    CStr(d(i, cols("Name") - b)), CStr(d(i, cols("DisplayID") - b)), _
                    CStr(d(i, cols("Status") - b)), CStr(d(i, cols("RowHash") - b)))
            End If
        Next i
    End If
    Set MasterInfo = info
End Function

' RelID -> Array(RelCode, SourceID, TargetID, Status, Label); also "code|s|t" -> RelID for active rows
Private Function RelInfo(byKey As Object) As Object
    Dim d As Variant, i As Long, info As Object
    Set info = modCore.NewDict()
    d = modCore.TableData(modData.RelLO())
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If CStr(d(i, 1)) <> "" Then
                info(CStr(d(i, 1))) = Array(CStr(d(i, 2)), CStr(d(i, 3)), CStr(d(i, 4)), CStr(d(i, 11)), _
                                            CStr(modCore.Nz(d(i, 9), "")))
                If CStr(d(i, 11)) <> "Retired" Then byKey(CStr(d(i, 2)) & "|" & d(i, 3) & "|" & d(i, 4)) = CStr(d(i, 1))
            End If
        Next i
    End If
    Set RelInfo = info
End Function

Private Function DomainOf(ByVal typeCode As String) As String
    If modMeta.gTypes.Exists(typeCode) Then DomainOf = modMeta.gTypes(typeCode)("Domain")
End Function

Private Function TypeOrder() As Object
    Dim d As Object, k As Variant, i As Long
    Set d = modCore.NewDict()
    For Each k In modMeta.gTypes.Keys
        i = i + 1
        d(k) = i
    Next k
    Set TypeOrder = d
End Function

Private Sub SortStrings(a() As String)
    Dim i As Long, j As Long, t As String
    For i = LBound(a) + 1 To UBound(a)
        t = a(i)
        j = i - 1
        Do While j >= LBound(a)
            If StrComp(a(j), t, vbBinaryCompare) <= 0 Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = t
    Next i
End Sub

Private Function Hex32ToDouble(ByVal h As String) As Double
    Dim i As Long, v As Double
    h = LCase$(h)
    For i = 1 To Len(h)
        v = v * 16# + (InStr("0123456789abcdef", Mid$(h, i, 1)) - 1)
    Next i
    Hex32ToDouble = v
End Function

' Order-independent fingerprint of the entities on a page (mirrors drawio.fingerprint)
Private Function Fingerprint(ids As Object, info As Object) As String
    Dim k As Variant, total As Double, hi As Long, lo As Long
    For Each k In ids.Keys
        If info.Exists(k) Then
            total = total + Hex32ToDouble(modCore.HashText(CStr(k) & ":" & info(k)(4)))
            total = total - Fix(total / 4294967296#) * 4294967296#
        End If
    Next k
    hi = CLng(Fix(total / 65536#))
    lo = CLng(total - CDbl(hi) * 65536#)
    Fingerprint = LCase$(Right$("0000" & Hex$(hi), 4) & Right$("0000" & Hex$(lo), 4))
End Function

' The repository entities shown anywhere in a diagram file; Nothing when it cannot be read
Public Function DiagramEntityIDs(ByVal path As String) As Object
    Dim dom As Object, msg As String, ids As Object, node As Object
    Set dom = LoadDrawio(path, msg)
    If dom Is Nothing Then Exit Function
    Set ids = modCore.NewDict()
    For Each node In dom.selectNodes("//object[@ea_id!=''] | //UserObject[@ea_id!='']")
        ids(Attr(node, "ea_id")) = True
    Next node
    Set DiagramEntityIDs = ids
End Function

' ================================================================ cells
Private Function AddEntityCell(dom As Object, root As Object, ByVal id As String, rec As Variant, _
                               ByVal x As Long, ByVal y As Long) As String
    Dim o As Object, c As Object, g As Object, cid As String
    cid = NewCellID()
    Set o = dom.createElement("object")
    o.setAttribute "id", cid
    o.setAttribute "label", HtmlEsc(CStr(rec(1)))
    o.setAttribute "ea_type", rec(0)
    o.setAttribute "ea_id", id
    o.setAttribute "ea_display", rec(2)
    o.setAttribute "ea_name", rec(1)
    o.setAttribute "ea_model", CStr(modCore.Setting("ModelID"))
    Set c = dom.createElement("mxCell")
    c.setAttribute "style", EntityStyle(rec(0))
    c.setAttribute "vertex", "1"
    c.setAttribute "parent", "1"
    Set g = dom.createElement("mxGeometry")
    g.setAttribute "x", CStr(x): g.setAttribute "y", CStr(y)
    g.setAttribute "width", CStr(SHAPE_W): g.setAttribute "height", CStr(SHAPE_H)
    g.setAttribute "as", "geometry"
    c.appendChild g
    o.appendChild c
    root.appendChild o
    AddEntityCell = cid
End Function

Private Sub AddBandCell(dom As Object, root As Object, ByVal band As String, ByVal y As Long)
    Dim c As Object, g As Object
    Set c = dom.createElement("mxCell")
    c.setAttribute "id", NewCellID()
    c.setAttribute "value", band
    c.setAttribute "style", BAND_STYLE
    c.setAttribute "vertex", "1"
    c.setAttribute "parent", "1"
    Set g = dom.createElement("mxGeometry")
    g.setAttribute "x", CStr(LEFT_X): g.setAttribute "y", CStr(y)
    g.setAttribute "width", "300": g.setAttribute "height", "24"
    g.setAttribute "as", "geometry"
    c.appendChild g
    root.appendChild c
End Sub

Private Sub AddRelCell(dom As Object, root As Object, ByVal relID As String, rel As Variant, _
                       ByVal srcCell As String, ByVal tgtCell As String)
    Dim o As Object, c As Object, g As Object
    Set o = dom.createElement("object")
    o.setAttribute "id", NewCellID()
    o.setAttribute "label", rel(4)
    o.setAttribute "ea_rel_id", relID
    o.setAttribute "ea_reltype", rel(0)
    Set c = dom.createElement("mxCell")
    c.setAttribute "style", RelStyle(rel(0))
    c.setAttribute "edge", "1"
    c.setAttribute "parent", "1"
    c.setAttribute "source", srcCell
    c.setAttribute "target", tgtCell
    Set g = dom.createElement("mxGeometry")
    g.setAttribute "relative", "1"
    g.setAttribute "as", "geometry"
    c.appendChild g
    o.appendChild c
    root.appendChild o
End Sub

' Wraps a plain <mxCell> in an <object> so it can carry ea_ properties
Private Function ConvertToObject(dom As Object, node As Object) As Object
    Dim o As Object
    If IsTaggable(node) Then Set ConvertToObject = node: Exit Function
    Set o = dom.createElement("object")
    o.setAttribute "id", Attr(node, "id")
    o.setAttribute "label", Attr(node, "value")
    node.parentNode.insertBefore o, node
    node.removeAttribute "id"
    node.removeAttribute "value"
    o.appendChild node
    Set ConvertToObject = o
End Function

' Refreshes labels of tagged shapes and marks entities that are retired or gone with a red outline.
' Returns the number of shapes changed; ids receives the ea_ids found on the page.
Private Function RefreshPage(root As Object, info As Object, ids As Object, ByRef nMissing As Long) As Long
    Dim node As Object, nv As Variant, c As Object, id As String, st As String, rec As Variant, changed As Long, gone As Boolean
    For Each nv In Children(root)
        Set node = nv
        If IsTaggable(node) Then
            id = Attr(node, "ea_id")
            Set c = InnerCell(node)
            If id <> "" And Not c Is Nothing Then
                If Attr(c, "vertex") = "1" Then
                    ids(id) = True
                    gone = Not info.Exists(id)
                    If Not gone Then gone = (info(id)(3) = "Retired")
                    st = Attr(c, "style")
                    If gone Then
                        nMissing = nMissing + 1
                        If Attr(node, "ea_missing") <> "1" Then
                            node.setAttribute "ea_missing", "1"
                            st = SetStyleKey(st, "strokeColor", "#FF0000")
                            st = SetStyleKey(st, "strokeWidth", "3")
                            c.setAttribute "style", st
                            changed = changed + 1
                        End If
                    Else
                        rec = info(id)
                        If Attr(node, "ea_missing") = "1" Then
                            node.removeAttribute "ea_missing"
                            st = SetStyleKey(st, "strokeColor", "")
                            st = SetStyleKey(st, "strokeWidth", "")
                            c.setAttribute "style", st
                            changed = changed + 1
                        End If
                        If Attr(node, "ea_display") <> rec(2) Then
                            node.setAttribute "ea_display", rec(2)
                            changed = changed + 1
                        End If
                        ' a label that differs from ea_name was renamed on the diagram: keep it for the harvest
                        If CleanLabel(Attr(node, "label")) = Attr(node, "ea_name") Or Attr(node, "ea_name") = "" Then
                            If CleanLabel(Attr(node, "label")) <> rec(1) Or Attr(node, "ea_name") <> rec(1) Then
                                node.setAttribute "label", HtmlEsc(CStr(rec(1)))
                                node.setAttribute "ea_name", rec(1)
                                changed = changed + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next nv
    RefreshPage = changed
End Function

' ================================================================ generate
' Writes or extends one page. Existing shapes keep their positions; entities and relationships
' not yet on the page are added below the existing content.
Public Function GenerateDiagram(ByVal path As String, ByVal pageName As String, ids As Object, _
                                ByRef msg As String, ByRef nAdded As Long, ByRef nRelAdded As Long) As Boolean
    Dim dom As Object, pg As Object, root As Object, info As Object, rels As Object, byKey As Object
    Dim onPage As Object, relOn As Object, node As Object, c As Object, g As Object, maxY As Double, y As Double
    Dim band As Variant, keys() As String, n As Long, k As Variant, i As Long, x As Long, y0 As Long
    Dim order As Object, bands As Collection, dom2 As String, nMiss As Long, pageIds As Object, done As Object
    Dim edgePairs As Object, cellEnt As Object, pairEnt As Object, pk As Variant, pp As Variant
    modMeta.LoadMeta
    LoadStyles
    Set info = MasterInfo()
    Set byKey = modCore.NewDict()
    Set rels = RelInfo(byKey)
    If FileExists(path) Then
        Set dom = LoadDrawio(path, msg)
        If dom Is Nothing Then Exit Function
    Else
        Set dom = NewMxfile()
    End If
    Set pg = FindPage(dom, pageName)
    If pg Is Nothing Then Set pg = AddPage(dom, pageName)
    Set root = PageRoot(pg)
    If root Is Nothing Then msg = CompressedMsg(pageName): Exit Function

    Set pageIds = modCore.NewDict()
    RefreshPage root, info, pageIds, nMiss
    ' what is already on the page
    Set onPage = modCore.NewDict(): Set relOn = modCore.NewDict()
    Set edgePairs = modCore.NewDict(): Set cellEnt = modCore.NewDict()
    For Each node In root.childNodes
        Set c = Nothing
        If node.nodeName = "mxCell" Or IsTaggable(node) Then Set c = InnerCell(node)
        If Not c Is Nothing Then
            If Attr(c, "vertex") = "1" Then
                If Attr(node, "ea_id") <> "" Then
                    If Not onPage.Exists(Attr(node, "ea_id")) Then onPage(Attr(node, "ea_id")) = Attr(node, "id")
                    cellEnt(Attr(node, "id")) = Attr(node, "ea_id")
                End If
                If Attr(c, "parent") = "1" Then
                    Set g = c.selectSingleNode("mxGeometry")
                    If Not g Is Nothing Then
                        y = Val(Attr(g, "y")) + Val(Attr(g, "height"))
                        If y > maxY Then maxY = y
                    End If
                End If
            ElseIf Attr(c, "edge") = "1" Then
                If Attr(node, "ea_rel_id") <> "" Then relOn(Attr(node, "ea_rel_id")) = True
                If Attr(node, "ea_rel_id") = "" Then edgePairs(Attr(c, "source") & "|" & Attr(c, "target")) = True
            End If
        End If
    Next node

    ' new entities, banded by domain, then type order and name
    Set order = TypeOrder()
    Set bands = New Collection
    For Each band In Split(BAND_ORDER, "|"): bands.Add CStr(band): Next band
    For Each k In ids.Keys
        If info.Exists(k) Then
            dom2 = DomainOf(info(k)(0))
            If InStr("|" & BAND_ORDER & "|", "|" & dom2 & "|") = 0 Then
                On Error Resume Next
                bands.Add dom2, dom2
                On Error GoTo 0
            End If
        End If
    Next k
    y0 = IIf(maxY > 0, CLng(maxY) + 60, TOP_Y)
    Set done = modCore.NewDict()
    For Each band In bands
        If Not done.Exists(band) Then
            done(band) = True
            n = 0
            ReDim keys(1 To ids.Count + 1)
            For Each k In ids.Keys
                If info.Exists(k) And Not onPage.Exists(k) Then
                    If info(k)(3) <> "Retired" And DomainOf(info(k)(0)) = band Then
                        n = n + 1
                        keys(n) = Format$(order(info(k)(0)), "0000") & "|" & LCase$(info(k)(1)) & "|" & k
                    End If
                End If
            Next k
            If n > 0 Then
                ReDim Preserve keys(1 To n)
                SortStrings keys
                AddBandCell dom, root, CStr(band), y0
                For i = 1 To n
                    k = Mid$(keys(i), InStrRev(keys(i), "|") + 1)
                    x = LEFT_X + ((i - 1) Mod GRID_COLS) * (SHAPE_W + GAP_X)
                    onPage(k) = AddEntityCell(dom, root, CStr(k), info(k), x, _
                                              y0 + 30 + ((i - 1) \ GRID_COLS) * (SHAPE_H + GAP_Y))
                    nAdded = nAdded + 1
                Next i
                y0 = y0 + 30 + ((n + GRID_COLS - 1) \ GRID_COLS) * (SHAPE_H + GAP_Y) + 20
            End If
        End If
    Next band

    ' relationships between entities on the page (skipping pairs already joined by a connector)
    Set pairEnt = modCore.NewDict()
    For Each pk In edgePairs.Keys
        pp = Split(pk, "|")
        If UBound(pp) = 1 Then
            If cellEnt.Exists(pp(0)) And cellEnt.Exists(pp(1)) Then pairEnt(cellEnt(pp(0)) & "|" & cellEnt(pp(1))) = True
        End If
    Next pk
    For Each k In rels.Keys
        If rels(k)(3) <> "Retired" And Not relOn.Exists(k) And Not pairEnt.Exists(rels(k)(1) & "|" & rels(k)(2)) Then
            If onPage.Exists(rels(k)(1)) And onPage.Exists(rels(k)(2)) Then
                AddRelCell dom, root, CStr(k), rels(k), onPage(rels(k)(1)), onPage(rels(k)(2))
                nRelAdded = nRelAdded + 1
            End If
        End If
    Next k

    SaveDom dom, path
    Set pageIds = modCore.NewDict()
    For Each k In onPage.Keys: pageIds(k) = True: Next k
    RegisterView path, pageName, pageIds, CountRels(root), info, "Generated"
    GenerateDiagram = True
End Function

Private Function CountRels(root As Object) As Long
    Dim node As Object
    For Each node In root.childNodes
        If IsTaggable(node) Then
            If Attr(node, "ea_rel_id") <> "" Then CountRels = CountRels + 1
        End If
    Next node
End Function

' ================================================================ form
Public Sub ShowDiagrams()
    Dim ws As Worksheet, core As String
    modCore.Heal
    modMeta.LoadMeta
    ThisWorkbook.Activate
    Set ws = ThisWorkbook.Worksheets(SH_FORM)
    core = "'" & ThisWorkbook.Name & "'!modDiagram."
    modCore.BeginBusy
    On Error GoTo Done
    ws.Activate
    modForms.ResetForm ws, "Diagrams (draw.io)", _
        "Diagrams live in Diagrams\ and the draw.io libraries in Library\. Close a diagram in draw.io before harvesting it."
    Section ws, 4, "1. Generate a diagram, or add to one"
    modForms.Field ws, 5, "Diagram file name *", "_file", "Overview"
    modForms.Field ws, 6, "Page name", "_page", "Page-1"
    modForms.Field ws, 7, "Content *", "_scope", "Around one entity"
    modForms.ListDropdown ws, "_scope", "DiagramScope"
    modForms.Field ws, 8, "Entity type (for one type)", "_etype"
    modForms.ListDropdown ws, "_etype", "TypeCodes"
    modForms.Field ws, 9, "Start entity Display ID (for around one entity)", "_start"
    modForms.Field ws, 10, "Relationship hops (1-5)", "_hops", 1
    ws.Cells(5, 4).Value = "Saved as Diagrams\<name>.drawio"
    ws.Cells(7, 4).Value = "If the page exists, shapes keep their places; missing entities and connectors are added below"
    modCore.AddButton ws, "Generate / add to diagram", core & "GenerateFromForm", ws.Range("F5"), 190

    Section ws, 12, "2. Harvest diagrams into the repository"
    ws.Cells(13, 1).Value = "New shapes, renamed shapes and new connectors are listed for review; nothing is written until you commit."
    ws.Cells(13, 1).Font.Color = RGB(89, 89, 89)
    modCore.AddButton ws, "Choose diagrams and preview", core & "HarvestPreview", ws.Range("F12"), 190
    modCore.AddButton ws, "Commit harvest", core & "HarvestCommit", ws.Range("F14"), 190

    Section ws, 16, "3. Keep diagrams current"
    ws.Cells(17, 1).Value = "Update refreshes names on every diagram and outlines retired or deleted entities in red."
    ws.Cells(17, 1).Font.Color = RGB(89, 89, 89)
    modCore.AddButton ws, "Update diagrams from model", core & "UpdateDiagrams", ws.Range("F16"), 190
    modCore.AddButton ws, "Views register", core & "ShowViews", ws.Range("F18"), 190

    Section ws, 20, "4. draw.io libraries"
    ws.Cells(21, 1).Value = "In draw.io: File > Open Library from > Device, then pick Library\ArchiMate4.xml and Library\ModelLibrary.xml."
    ws.Cells(21, 1).Font.Color = RGB(89, 89, 89)
    modCore.AddButton ws, "Rebuild ArchiMate 4 stencil", core & "BuildStencilButton", ws.Range("F20"), 190
    modCore.AddButton ws, "Rebuild Model Library", core & "BuildModelLibraryButton", ws.Range("F22"), 190

    Section ws, 24, "5. Open and export"
    modForms.Field ws, 25, "Diagram to open (file name)", "_open", "Overview"
    modForms.Field ws, 26, "Image format", "_imgfmt", "png"
    modForms.ListDropdown ws, "_imgfmt", "ImageFormat"
    ws.Cells(26, 4).Value = "Images of every page go to Diagrams\Export\"
    modCore.AddButton ws, "Open diagram", core & "OpenFromForm", ws.Range("F25"), 190
    modCore.AddButton ws, "Export images", core & "ExportImagesFromForm", ws.Range("F27"), 190
    modForms.Field ws, 29, "draw.io program", "_exe", DrawioExe()
    ws.Cells(29, 4).Value = "Found automatically; set DrawioPath in SYS_Settings if draw.io is installed elsewhere"
    ws.Range("B29").Locked = True
    modCore.AddButton ws, "Back to Home", "'" & ThisWorkbook.Name & "'!modForms.GoHome", ws.Range("F31"), 190
    ws.Range("B5").Select
Done:
    Dim eMsg As String
    If Err.Number <> 0 Then eMsg = Err.Description
    On Error Resume Next
    If eMsg <> "" Then modForms.FormMessage ws, "Form not fully built: " & eMsg, True
    modCore.ProtectSheet ws
    modCore.EndBusy
End Sub

Private Sub Section(ws As Worksheet, ByVal r As Long, ByVal text As String)
    ws.Cells(r, 1).Value = text
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    ws.Cells(r, 1).Font.Color = RGB(31, 56, 100)
End Sub

Private Function DiagramPath(ByVal nm As String) As String
    nm = SafeFile(nm)
    If nm = "" Then Exit Function
    If LCase$(Right$(nm, 7)) <> ".drawio" Then nm = nm & ".drawio"
    DiagramPath = DiagramsFolder() & nm
End Function

Public Sub GenerateFromForm()
    Dim ws As Worksheet, path As String, pageName As String, scope As String, ids As Object, msg As String
    Dim nAdd As Long, nRel As Long
    modCore.Heal
    Set ws = ThisWorkbook.Worksheets(SH_FORM)
    FormMsg ws, "", False
    path = DiagramPath(CStr(modForms.FieldValue(ws, "_file")))
    pageName = Trim$(CStr(modForms.FieldValue(ws, "_page")))
    If pageName = "" Then pageName = "Page-1"
    scope = CStr(modForms.FieldValue(ws, "_scope"))
    If path = "" Then FormMsg ws, "Enter a diagram file name.", True: Exit Sub
    Set ids = modIO.SelectEntities(scope, CStr(modForms.FieldValue(ws, "_etype")), _
                                   CStr(modForms.FieldValue(ws, "_start")), _
                                   Val(CStr(modForms.FieldValue(ws, "_hops"))), False, msg)
    If msg <> "" Then FormMsg ws, msg, True: Exit Sub
    If ids.Count = 0 Then FormMsg ws, "Nothing to draw for that choice.", True: Exit Sub
    If ids.Count > 300 Then
        If MsgBox(ids.Count & " entities is a large diagram. Draw it anyway?", vbYesNo + vbQuestion, _
                  "EA Workbench") = vbNo Then Exit Sub
    End If
    RunGenerate path, pageName, ids
End Sub

Private Sub RunGenerate(ByVal path As String, ByVal pageName As String, ids As Object)
    Dim msg As String, nAdd As Long, nRel As Long
    On Error GoTo Fail
    modCore.BeginBusy
    modCore.Status "drawing " & ids.Count & " entities"
    If Not GenerateDiagram(path, pageName, ids, msg, nAdd, nRel) Then
        modCore.EndBusy
        MsgBox "Diagram not written: " & msg, vbExclamation, "EA Workbench"
        Exit Sub
    End If
    modSync.SaveWithStamp ThisWorkbook
    modCore.EndBusy
    If MsgBox("Diagram saved: " & path & vbLf & nAdd & " shape(s) and " & nRel & " connector(s) added to page """ & _
              pageName & """." & vbLf & vbLf & "Open it in draw.io now?", vbYesNo + vbInformation, "EA Workbench") = vbYes Then
        OpenDiagram path
    End If
    Exit Sub
Fail:
    Dim m As String
    m = Err.Description
    modCore.ResetBusy
    MsgBox "Diagram not written: " & m, vbExclamation, "EA Workbench"
End Sub

' Right-click on a type sheet: draw the selected entities and the relationships between them
Public Sub DiagramOfSelected()
    Dim c As Collection, ids As Object, k As Variant, nm As String, path As String
    modCore.Heal
    modMeta.LoadMeta
    Set c = modReports.ChosenEntities("Display IDs to draw, separated by commas (e.g. APP-0001, CAP-0001):")
    If c.Count = 0 Then Exit Sub
    Set ids = modCore.NewDict()
    For Each k In c: ids(k) = True: Next k
    If c.Count = 1 Then
        Set ids = modIO.Neighbourhood(CStr(c(1)), 1)
        nm = modData.EntityNameOf(CStr(c(1)))
    Else
        nm = "Selection"
    End If
    nm = InputBox("Diagram file name (in Diagrams\). An existing file gets a new page.", "EA Workbench", SafeFile(nm))
    If nm = "" Then Exit Sub
    path = DiagramPath(nm)
    RunGenerate path, "Page-" & Format$(Now, "yyyymmdd-hhnn"), ids
End Sub

Private Sub FormMsg(ws As Worksheet, ByVal text As String, ByVal isError As Boolean)
    modCore.UnprotectSheet ws
    modForms.FormMessage ws, text, isError
    modCore.ProtectSheet ws
End Sub

' ================================================================ update diagrams
Public Sub UpdateDiagrams()
    Dim f As String, path As String, dom As Object, pg As Object, root As Object, info As Object, msg As String
    Dim nFiles As Long, nShapes As Long, nMiss As Long, changed As Long, ids As Object, problems As String
    Dim fv As Variant
    modCore.Heal
    On Error GoTo Fail
    modMeta.LoadMeta
    LoadStyles
    modCore.BeginBusy
    Set info = MasterInfo()
    For Each fv In DiagramFiles()
        f = CStr(fv)
        path = DiagramsFolder() & f
        msg = ""
        Set dom = LoadDrawio(path, msg)
        If dom Is Nothing Then
            problems = problems & vbLf & f & ": " & msg
        Else
            changed = 0
            For Each pg In dom.documentElement.selectNodes("diagram")
                Set root = PageRoot(pg)
                If root Is Nothing Then
                    problems = problems & vbLf & f & ": " & CompressedMsg(Attr(pg, "name"))
                Else
                    Set ids = modCore.NewDict()
                    changed = changed + RefreshPage(root, info, ids, nMiss)
                    RegisterView path, Attr(pg, "name"), ids, CountRels(root), info, "Updated"
                End If
            Next pg
            If changed > 0 Then SaveDom dom, path: nShapes = nShapes + changed
            nFiles = nFiles + 1
        End If
    Next fv
    modSync.SaveWithStamp ThisWorkbook
    modCore.EndBusy
    MsgBox nFiles & " diagram file(s) checked, " & nShapes & " shape(s) updated, " & nMiss & _
           " shape(s) show entities that are retired or deleted (red outline)." & _
           IIf(problems <> "", vbLf & vbLf & "Not updated:" & problems, "") & vbLf & vbLf & _
           "If a diagram is open in draw.io, close it without saving and reopen it to see the changes.", _
           IIf(problems <> "", vbExclamation, vbInformation), "EA Workbench"
    Exit Sub
Fail:
    Dim m As String
    m = Err.Description
    modCore.ResetBusy
    MsgBox "Update stopped: " & m, vbExclamation, "EA Workbench"
End Sub

' ================================================================ libraries
Public Sub BuildStencilButton()
    Dim p As String
    modCore.Heal
    p = BuildStencil()
    If p <> "" Then MsgBox "ArchiMate 4 stencil written to " & p & vbLf & vbLf & _
        "In draw.io: File > Open Library from > Device, and pick that file. If it is already open, close the library " & _
        "(right-click its title) and open it again.", vbInformation, "EA Workbench"
End Sub

Public Sub BuildModelLibraryButton()
    Dim p As String, n As Long
    modCore.Heal
    p = BuildModelLibrary(n)
    If p <> "" Then MsgBox n & " entities written to " & p & vbLf & vbLf & _
        "In draw.io: File > Open Library from > Device, and pick that file. Drag an entity from it to put the real " & _
        "repository object on a diagram.", vbInformation, "EA Workbench"
End Sub

Private Function ShapeModel(ByVal label As String, attrs As Collection, ByVal style As String, ByVal w As Long, _
                            ByVal h As Long, ByVal isEdge As Boolean) As String
    Dim s As String, kv As Variant
    s = "<mxGraphModel><root><mxCell id=""0""/><mxCell id=""1"" parent=""0""/><object id=""2"" label=""" & XmlEsc(HtmlEsc(label)) & """"
    For Each kv In attrs
        s = s & " " & kv(0) & "=""" & XmlEsc(CStr(kv(1))) & """"
    Next kv
    s = s & ">"
    If isEdge Then
        s = s & "<mxCell style=""" & XmlEsc(style) & """ edge=""1"" parent=""1""><mxGeometry width=""" & w & _
            """ relative=""1"" as=""geometry""><mxPoint x=""0"" y=""0"" as=""sourcePoint""/><mxPoint x=""" & w & _
            """ y=""0"" as=""targetPoint""/></mxGeometry></mxCell>"
    Else
        s = s & "<mxCell style=""" & XmlEsc(style) & """ vertex=""1"" parent=""1""><mxGeometry width=""" & w & _
            """ height=""" & h & """ as=""geometry""/></mxCell>"
    End If
    ShapeModel = s & "</object></root></mxGraphModel>"
End Function

Private Function LibEntry(ByVal xml As String, ByVal w As Long, ByVal h As Long, ByVal title As String) As String
    LibEntry = "{""xml"":""" & Compress(xml) & """,""w"":" & w & ",""h"":" & h & ",""aspect"":""fixed"",""title"":" & _
               JsonStr(title) & "}"
End Function

Private Sub WriteLibrary(ByVal path As String, entries As Collection)
    Dim dom As Object, root As Object, s As String, e As Variant, parts() As String, i As Long
    If entries.Count > 0 Then
        ReDim parts(1 To entries.Count)
        For Each e In entries
            i = i + 1
            parts(i) = e
        Next e
        s = "[" & Join(parts, ",") & "]"
    Else
        s = "[]"
    End If
    Set dom = NewDom()
    Set root = dom.createElement("mxlibrary")
    root.Text = s
    dom.appendChild root
    dom.Save path
End Sub

Private Function KV(ByVal k As String, ByVal v As String) As Variant
    KV = Array(k, v)
End Function

' Library\ArchiMate4.xml: one tagged shape per active entity type (with its mandatory attributes as
' blank properties for Edit Data), then one connector per active relationship type.
Public Function BuildStencil(Optional ByVal quiet As Boolean = False) As String
    Dim entries As New Collection, t As Variant, rc As Variant, attrs As Collection, a As Variant, path As String
    On Error GoTo Fail
    modMeta.LoadMeta
    LoadStyles
    For Each t In modMeta.gTypes.Keys
        If modMeta.gTypes(t)("Status") = "Active" Then
            Set attrs = New Collection
            attrs.Add KV("ea_type", CStr(t)): attrs.Add KV("ea_id", ""): attrs.Add KV("ea_model", CStr(modCore.Setting("ModelID")))
            For Each a In modMeta.AttrsForType(CStr(t))
                If modMeta.IsMandatory(CStr(t), CStr(a)) Then attrs.Add KV(CStr(a), "")
            Next a
            entries.Add LibEntry(ShapeModel(modMeta.TypeName_(CStr(t)), attrs, EntityStyle(CStr(t)), SHAPE_W, SHAPE_H, False), _
                                 SHAPE_W, SHAPE_H, modMeta.TypeName_(CStr(t)))
        End If
    Next t
    For Each rc In modMeta.gRelTypes.Keys
        If modMeta.gRelTypes(rc)("Status") = "Active" Then
            Set attrs = New Collection
            attrs.Add KV("ea_reltype", CStr(rc)): attrs.Add KV("ea_rel_id", "")
            entries.Add LibEntry(ShapeModel("", attrs, RelStyle(CStr(rc)), SHAPE_W, 0, True), SHAPE_W, 0, _
                                 CStr(rc) & " (" & modMeta.RelVerb(CStr(rc)) & ")")
        End If
    Next rc
    path = SubFolder("Library") & "ArchiMate4.xml"
    WriteLibrary path, entries
    BuildStencil = path
    Exit Function
Fail:
    If Not quiet Then MsgBox "Stencil not written: " & Err.Description, vbExclamation, "EA Workbench"
End Function

' Library\ModelLibrary.xml: every active entity as a ready-tagged shape, by type then name
Public Function BuildModelLibrary(Optional ByRef n As Long, Optional ByVal quiet As Boolean = False) As String
    Dim entries As New Collection, info As Object, order As Object, keys() As String, k As Variant, i As Long
    Dim attrs As Collection, rec As Variant, id As String, path As String
    On Error GoTo Fail
    modMeta.LoadMeta
    LoadStyles
    Set info = MasterInfo()
    Set order = TypeOrder()
    n = 0
    If info.Count > 0 Then
        ReDim keys(1 To info.Count)
        For Each k In info.Keys
            If info(k)(3) <> "Retired" And order.Exists(info(k)(0)) Then
                n = n + 1
                keys(n) = Format$(order(info(k)(0)), "0000") & "|" & LCase$(info(k)(1)) & "|" & k
            End If
        Next k
    End If
    If n > 0 Then
        ReDim Preserve keys(1 To n)
        SortStrings keys
        For i = 1 To n
            id = Mid$(keys(i), InStrRev(keys(i), "|") + 1)
            rec = info(id)
            Set attrs = New Collection
            attrs.Add KV("ea_type", rec(0)): attrs.Add KV("ea_id", id): attrs.Add KV("ea_display", rec(2))
            attrs.Add KV("ea_name", rec(1)): attrs.Add KV("ea_model", CStr(modCore.Setting("ModelID")))
            entries.Add LibEntry(ShapeModel(CStr(rec(1)), attrs, EntityStyle(CStr(rec(0))), SHAPE_W, SHAPE_H, False), _
                                 SHAPE_W, SHAPE_H, modMeta.TypeName_(CStr(rec(0))) & ": " & rec(1) & " (" & rec(2) & ")")
        Next i
    End If
    path = SubFolder("Library") & "ModelLibrary.xml"
    WriteLibrary path, entries
    BuildModelLibrary = path
    Exit Function
Fail:
    If Not quiet Then MsgBox "Model Library not written: " & Err.Description, vbExclamation, "EA Workbench"
End Function

' Called at the end of a full sync: refresh the Model Library and the views register, quietly
Public Sub AfterFullSync()
    On Error Resume Next
    BuildStencil True
    BuildModelLibrary 0, True
    CheckViews True
End Sub

' ================================================================ views register
Private Function ViewsLO() As ListObject
    Set ViewsLO = modCore.Tbl(SH_VIEWS, "tblViews")
End Function

Private Sub RegisterView(ByVal path As String, ByVal pageName As String, ids As Object, ByVal nRels As Long, _
                         info As Object, ByVal action As String)
    Dim lo As ListObject, d As Variant, i As Long, r As Long, vf As String, ws As Worksheet, row As Range
    Dim nMiss As Long, k As Variant
    Set lo = ViewsLO()
    Set ws = lo.Parent
    vf = ViewFile(path)
    For Each k In ids.Keys
        If Not info.Exists(k) Then
            nMiss = nMiss + 1
        ElseIf info(k)(3) = "Retired" Then
            nMiss = nMiss + 1
        End If
    Next k
    d = modCore.TableData(lo)
    If IsArray(d) Then
        For i = 1 To UBound(d, 1)
            If StrComp(CStr(d(i, 1)), vf, vbTextCompare) = 0 And StrComp(CStr(d(i, 2)), pageName, vbTextCompare) = 0 Then
                r = i: Exit For
            End If
        Next i
    End If
    modCore.UnprotectSheet ws
    If r = 0 Then
        Set row = modCore.AppendRow(lo, Array(vf, pageName, 0, 0, "", "", "", "", ""))
    Else
        Set row = lo.ListRows(r).Range
    End If
    row.Cells(1, 3).Value = ids.Count
    row.Cells(1, 4).Value = nRels
    If action = "Generated" Then row.Cells(1, 5).Value = modCore.UtcNow()
    If action = "Harvested" Then row.Cells(1, 6).Value = modCore.UtcNow()
    row.Cells(1, 7).NumberFormat = "@"
    row.Cells(1, 7).Value = Fingerprint(ids, info)
    row.Cells(1, 8).Value = IIf(nMiss > 0, "Stale: " & nMiss & " entit" & IIf(nMiss = 1, "y", "ies") & _
                                " retired or deleted", "OK")
    row.Cells(1, 9).Value = action & " " & modCore.UtcNow()
    modCore.ProtectSheet ws
End Sub

' Re-reads every registered page; returns the number that are not OK
Public Function CheckViews(Optional ByVal quiet As Boolean = False) As Long
    Dim lo As ListObject, d As Variant, i As Long, path As String, dom As Object, pg As Object, root As Object
    Dim info As Object, ids As Object, node As Object, c As Object, st As String, msg As String, nMiss As Long
    Dim rels As Object, byKey As Object, k As Variant, nNot As Long, fp As String, ws As Worksheet, cache As Object
    Set lo = ViewsLO()
    Set ws = lo.Parent
    d = modCore.TableData(lo)
    If Not IsArray(d) Then Exit Function
    modMeta.LoadMeta
    Set info = MasterInfo()
    Set byKey = modCore.NewDict()
    Set rels = RelInfo(byKey)
    Set cache = modCore.NewDict()
    modCore.UnprotectSheet ws
    For i = 1 To UBound(d, 1)
        If CStr(d(i, 1)) <> "" Then
            path = FullPath(CStr(d(i, 1)))
            st = ""
            If Not FileExists(path) Then
                st = "File missing"
            Else
                If cache.Exists(LCase$(path)) Then
                    Set dom = cache(LCase$(path))
                Else
                    msg = ""
                    Set dom = LoadDrawio(path, msg)
                    If Not dom Is Nothing Then Set cache(LCase$(path)) = dom
                End If
                If dom Is Nothing Then
                    st = "Unreadable: " & msg
                Else
                    Set pg = FindPage(dom, CStr(d(i, 2)))
                    If pg Is Nothing Then
                        st = "Page missing"
                    Else
                        Set root = PageRoot(pg)
                        If root Is Nothing Then
                            st = "Compressed - cannot read"
                        Else
                            Set ids = modCore.NewDict()
                            For Each node In root.childNodes
                                If IsTaggable(node) Then
                                    If Attr(node, "ea_id") <> "" Then ids(Attr(node, "ea_id")) = True
                                End If
                            Next node
                            nMiss = 0
                            For Each k In ids.Keys
                                If Not info.Exists(k) Then
                                    nMiss = nMiss + 1
                                ElseIf info(k)(3) = "Retired" Then
                                    nMiss = nMiss + 1
                                End If
                            Next k
                            nNot = 0
                            For Each k In byKey.Keys
                                If ids.Exists(rels(byKey(k))(1)) And ids.Exists(rels(byKey(k))(2)) Then
                                    If Not RelOnPage(root, CStr(byKey(k))) Then nNot = nNot + 1
                                End If
                            Next k
                            fp = Fingerprint(ids, info)
                            If nMiss > 0 Then
                                st = "Stale: " & nMiss & " entit" & IIf(nMiss = 1, "y", "ies") & " retired or deleted"
                            ElseIf fp <> CStr(d(i, 7)) Then
                                st = "Stale: entities changed since last generate, update or harvest"
                            Else
                                st = "OK"
                            End If
                            If nNot > 0 Then st = st & "; " & nNot & " relationship(s) not drawn"
                            lo.ListRows(i).Range.Cells(1, 3).Value = ids.Count
                        End If
                    End If
                End If
            End If
            lo.ListRows(i).Range.Cells(1, 8).Value = st
            If st <> "OK" Then CheckViews = CheckViews + 1
        End If
    Next i
    modCore.ProtectSheet ws
End Function

Private Function RelOnPage(root As Object, ByVal relID As String) As Boolean
    RelOnPage = Not root.selectSingleNode("*[@ea_rel_id='" & Replace(relID, "'", "") & "']") Is Nothing
End Function

Public Sub ShowViews()
    Dim ws As Worksheet, n As Long, core As String
    modCore.Heal
    n = CheckViews(False)
    Set ws = ThisWorkbook.Worksheets(SH_VIEWS)
    core = "'" & ThisWorkbook.Name & "'!modDiagram."
    ThisWorkbook.Activate
    modCore.BeginBusy
    modCore.UnprotectSheet ws
    modCore.DeleteButtons ws
    ws.Range("A1").Value = "Views register"
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = n & " page(s) need attention. Stale pages: run Update diagrams from model, or generate the page again " & _
                           "to add missing entities and connectors."
    ws.Range("A2").Font.Italic = True
    modCore.AddButton ws, "Open selected diagram", core & "OpenSelectedView", ws.Range("D1"), 170
    modCore.AddButton ws, "Update diagrams from model", core & "UpdateDiagrams", ws.Range("F1"), 190
    modCore.AddButton ws, "Back to Diagrams", core & "ShowDiagrams", ws.Range("H1"), 150
    modCore.ProtectSheet ws
    modCore.EndBusy
    ws.Activate
End Sub

Public Sub OpenSelectedView()
    Dim lo As ListObject, r As Long
    Set lo = ViewsLO()
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If ActiveSheet.Name <> SH_VIEWS Then Exit Sub
    r = ActiveCell.Row - lo.HeaderRowRange.Row
    If r < 1 Or r > lo.ListRows.Count Then
        MsgBox "Click a row of the register first.", vbInformation, "EA Workbench"
        Exit Sub
    End If
    OpenDiagram FullPath(CStr(lo.ListRows(r).Range.Cells(1, 1).Value))
End Sub

' ================================================================ open and export
Public Function DrawioExe() As String
    Dim c As Variant
    For Each c In Array(CStr(modCore.Setting("DrawioPath")), _
                        Environ$("LOCALAPPDATA") & "\Programs\draw.io\draw.io.exe", _
                        Environ$("ProgramFiles") & "\draw.io\draw.io.exe", _
                        Environ$("ProgramFiles(x86)") & "\draw.io\draw.io.exe")
        If CStr(c) <> "" Then
            If FileExists(CStr(c)) Then DrawioExe = CStr(c): Exit Function
        End If
    Next c
End Function

Public Sub OpenDiagram(ByVal path As String)
    Dim exe As String
    On Error GoTo Fail
    If Not FileExists(path) Then
        MsgBox "Diagram not found: " & path, vbExclamation, "EA Workbench"
        Exit Sub
    End If
    exe = DrawioExe()
    If exe <> "" Then
        Shell """" & exe & """ """ & path & """", vbNormalFocus
    Else
        ThisWorkbook.FollowHyperlink path
    End If
    Exit Sub
Fail:
    MsgBox "Could not open the diagram: " & Err.Description & vbLf & _
           "Check DrawioPath in SYS_Settings, or open the file from draw.io.", vbExclamation, "EA Workbench"
End Sub

Public Sub OpenFromForm()
    Dim ws As Worksheet, p As String, f As Variant
    modCore.Heal
    Set ws = ThisWorkbook.Worksheets(SH_FORM)
    p = DiagramPath(CStr(modForms.FieldValue(ws, "_open")))
    If p = "" Or Not FileExists(p) Then
        On Error Resume Next
        ChDrive Left$(DiagramsFolder(), 1)
        ChDir DiagramsFolder()
        On Error GoTo 0
        f = Application.GetOpenFilename("draw.io diagrams (*.drawio),*.drawio", , "Open diagram")
        If VarType(f) = vbBoolean Then Exit Sub
        p = CStr(f)
    End If
    OpenDiagram p
End Sub

Public Sub ExportImagesFromForm()
    Dim ws As Worksheet, fmt As String
    modCore.Heal
    Set ws = ThisWorkbook.Worksheets(SH_FORM)
    fmt = LCase$(Trim$(CStr(modForms.FieldValue(ws, "_imgfmt"))))
    If fmt <> "png" And fmt <> "svg" And fmt <> "pdf" Then fmt = "png"
    ExportImages fmt
End Sub

' Exports every page of every diagram with the draw.io command line (draw.io.exe -x)
Public Sub ExportImages(ByVal fmt As String)
    Dim exe As String, f As String, path As String, outDir As String, dom As Object, pg As Object, i As Long
    Dim sh As Object, cmd As String, outFile As String, n As Long, msg As String, fails As String, rc As Long, base As String
    Dim fv As Variant, folder As String
    exe = DrawioExe()
    If exe = "" Then
        MsgBox "draw.io Desktop was not found. Install it, or set DrawioPath in SYS_Settings to draw.io.exe.", _
               vbExclamation, "EA Workbench"
        Exit Sub
    End If
    On Error GoTo Fail
    outDir = SubFolder("Diagrams\Export")
    Set sh = CreateObject("WScript.Shell")
    Application.Cursor = xlWait
    folder = DiagramsFolder()
    For Each fv In DiagramFiles()
        f = CStr(fv)
        path = folder & f
        base = Left$(f, Len(f) - 7)
        If fmt = "pdf" Then
            outFile = outDir & base & ".pdf"
            cmd = """" & exe & """ -x -f pdf -a -o """ & outFile & """ """ & path & """"
            KillQuietly outFile
            rc = sh.Run(cmd, 0, True)
            If FileExists(outFile) Then n = n + 1 Else fails = fails & vbLf & f
        Else
            msg = ""
            Set dom = LoadDrawio(path, msg)
            If dom Is Nothing Then
                fails = fails & vbLf & f & " (" & msg & ")"
            Else
                i = 0
                For Each pg In dom.documentElement.selectNodes("diagram")
                    outFile = outDir & base & "-" & SafeFile(Attr(pg, "name")) & "." & fmt
                    cmd = """" & exe & """ -x -f " & fmt & " -p " & (i + 1) & " -o """ & outFile & """ """ & path & """"
                    Application.StatusBar = "EA Workbench: exporting " & f & " page " & (i + 1)
                    KillQuietly outFile
                    rc = sh.Run(cmd, 0, True)
                    If FileExists(outFile) Then n = n + 1 Else fails = fails & vbLf & f & " page " & (i + 1)
                    i = i + 1
                Next pg
            End If
        End If
    Next fv
    Application.Cursor = xlDefault
    Application.StatusBar = False
    MsgBox n & " image(s) written to " & outDir & IIf(fails <> "", vbLf & vbLf & "Not exported:" & fails, ""), _
           IIf(fails <> "", vbExclamation, vbInformation), "EA Workbench"
    Exit Sub
Fail:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    MsgBox "Export stopped: " & Err.Description, vbExclamation, "EA Workbench"
End Sub

Private Sub KillQuietly(ByVal path As String)
    On Error Resume Next
    If FileExists(path) Then Kill path
End Sub

' ================================================================ harvest
Private Function ReviewLO() As ListObject
    Set ReviewLO = modCore.Tbl(SH_REVIEW, "tblHarvest")
End Function

Public Sub HarvestPreview()
    Dim files As Variant, list As String, i As Long
    modCore.Heal
    On Error Resume Next
    ChDrive Left$(DiagramsFolder(), 1)
    ChDir DiagramsFolder()
    On Error GoTo 0
    files = Application.GetOpenFilename("draw.io diagrams (*.drawio;*.xml),*.drawio;*.xml", , _
                                        "Choose diagrams to harvest (close them in draw.io first)", , True)
    If Not IsArray(files) Then Exit Sub
    For i = LBound(files) To UBound(files)
        list = list & IIf(list = "", "", "|") & files(i) & "*" & FileStamp(CStr(files(i)))
    Next i
    RunHarvest list, False
End Sub

Public Sub HarvestCommit()
    Dim list As String, ws As Worksheet, part As Variant, p() As String, changed As String
    modCore.Heal
    Set ws = ThisWorkbook.Worksheets(SH_REVIEW)
    list = CStr(ws.Range("B4").Value)
    If list = "" Then
        MsgBox "Preview a harvest first (Choose diagrams and preview).", vbInformation, "EA Workbench"
        Exit Sub
    End If
    For Each part In Split(list, "|")
        p = Split(part, "*")
        If UBound(p) >= 1 Then
            If FileStamp(p(0)) <> p(1) Then changed = changed & vbLf & p(0)
        End If
    Next part
    If changed <> "" Then
        MsgBox "These diagrams changed after the preview. Preview again before committing:" & changed, _
               vbExclamation, "EA Workbench"
        Exit Sub
    End If
    If MsgBox("Commit this harvest? Rows marked Error are skipped. A snapshot is taken first, and the new IDs are " & _
              "written into the diagram files." & vbLf & vbLf & "Close these diagrams in draw.io first.", _
              vbOKCancel + vbQuestion, "EA Workbench") = vbCancel Then Exit Sub
    If modEvents.PendingCount() > 0 Then modEvents.PushPending True
    If Not modHistory.AutoSnapshot("Before diagram harvest") Then Exit Sub
    RunHarvest list, True
End Sub

Private Sub RunHarvest(ByVal fileList As String, ByVal commit As Boolean)
    Dim rows As New Collection, batch As String, touched As Object, part As Variant, p() As String
    Dim info As Object, rels As Object, byKey As Object, nameIdx As Object, plan As Object
    Dim cnt(1 To 6) As Long   ' new, renamed, linked, new rels, errors, unchanged
    On Error GoTo Fail
    modCore.BeginBusy
    modMeta.LoadMeta True
    LoadStyles
    batch = modCore.NewGUID()
    Set touched = modCore.NewDict()
    Set info = MasterInfo()
    Set byKey = modCore.NewDict()
    Set rels = RelInfo(byKey)
    Set nameIdx = modData.BuildNameIndex()
    Set plan = modCore.NewDict()
    For Each part In Split(fileList, "|")
        p = Split(part, "*")
        HarvestFile p(0), commit, batch, rows, touched, info, rels, byKey, nameIdx, plan, cnt
    Next part
    WriteReview fileList, rows, commit
    If commit Then
        If touched.Count > 0 Then modPaint.PaintTypes touched
        modSync.SaveWithStamp modCore.EntitiesWB(False)
        modSync.SaveWithStamp ThisWorkbook
        If cnt(1) + cnt(2) > 0 Then BuildModelLibrary 0, True
        modCore.EndBusy
        MsgBox "Harvest committed." & vbLf & vbLf & cnt(1) & " new entit" & IIf(cnt(1) = 1, "y", "ies") & ", " & _
               cnt(2) & " renamed, " & cnt(3) & " shape(s) linked to existing entities, " & cnt(4) & _
               " new relationship(s)." & vbLf & cnt(5) & " item(s) skipped with errors - see HARVEST_Review." & vbLf & vbLf & _
               "The diagram files now carry the new IDs. Reopen them in draw.io.", _
               IIf(cnt(5) > 0, vbExclamation, vbInformation), "EA Workbench"
    Else
        modCore.EndBusy
        MsgBox "Preview ready: " & cnt(1) & " new entit" & IIf(cnt(1) = 1, "y", "ies") & ", " & cnt(2) & " rename(s), " & _
               cnt(3) & " link(s) to existing entities, " & cnt(4) & " new relationship(s), " & cnt(5) & " error(s)." & vbLf & _
               "Check HARVEST_Review, then click Commit harvest.", IIf(cnt(5) > 0, vbExclamation, vbInformation), "EA Workbench"
    End If
    Exit Sub
Fail:
    Dim msg As String
    msg = Err.Description
    modCore.ResetBusy
    MsgBox "Harvest stopped: " & msg, vbExclamation, "EA Workbench"
End Sub

Private Sub Row_(rows As Collection, ByVal f As String, ByVal pg As String, ByVal cellID As String, ByVal kind As String, _
                 ByVal action As String, ByVal id As String, ByVal disp As String, ByVal nm As String, ByVal tc As String, _
                 ByVal field As String, ByVal curV As Variant, ByVal newV As Variant, ByVal sev As String, ByVal msg As String)
    rows.Add Array(f, pg, cellID, kind, action, id, disp, nm, tc, field, curV, newV, sev, msg)
End Sub

Private Sub IndexAdd(nameIdx As Object, ByVal typeCode As String, ByVal nm As String, ByVal id As String)
    Dim k As Variant
    For Each k In Array(typeCode & "|" & LCase$(Trim$(nm)), "*|" & LCase$(Trim$(nm)))
        If Not nameIdx.Exists(k) Then nameIdx.Add k, modCore.NewDict()
        nameIdx(k)(id) = True
    Next k
End Sub

Private Sub IndexRemove(nameIdx As Object, ByVal typeCode As String, ByVal nm As String, ByVal id As String)
    Dim k As Variant
    For Each k In Array(typeCode & "|" & LCase$(Trim$(nm)), "*|" & LCase$(Trim$(nm)))
        If nameIdx.Exists(k) Then
            If nameIdx(k).Exists(id) Then nameIdx(k).Remove id
            If nameIdx(k).Count = 0 Then nameIdx.Remove k
        End If
    Next k
End Sub

Private Function FirstKey(d As Object) As String
    Dim k As Variant
    For Each k In d.Keys
        FirstKey = CStr(k): Exit Function
    Next k
End Function

Private Sub HarvestFile(ByVal path As String, ByVal commit As Boolean, ByVal batch As String, rows As Collection, _
                        touched As Object, info As Object, rels As Object, byKey As Object, nameIdx As Object, _
                        plan As Object, cnt() As Long)
    Dim dom As Object, msg As String, f As String, pg As Object, root As Object, pageName As String, dirty As Boolean
    Dim node As Object, nv As Variant, c As Object, cellOf As Object, typeOf As Object, src As String, ids As Object, nRels As Long
    f = Mid$(path, InStrRev(path, Application.PathSeparator) + 1)
    src = "Diagram " & f
    Set dom = LoadDrawio(path, msg)
    If dom Is Nothing Then
        Row_ rows, f, "", "", "File", "Error", "", "", "", "", "", "", "", "Error", msg
        cnt(5) = cnt(5) + 1
        Exit Sub
    End If
    For Each pg In dom.documentElement.selectNodes("diagram")
        pageName = Attr(pg, "name")
        Set root = PageRoot(pg)
        If root Is Nothing Then
            Row_ rows, f, pageName, "", "Page", "Error", "", "", "", "", "", "", "", "Error", CompressedMsg(pageName)
            cnt(5) = cnt(5) + 1
        Else
            Set cellOf = modCore.NewDict()     ' cell id -> EntityID or "(new)|type|name" (preview) ; "" = unusable
            Set typeOf = modCore.NewDict()     ' cell id -> TypeCode
            ' ---- shapes
            For Each nv In Children(root)
                Set node = nv
                Set c = Nothing
                If node.nodeName = "mxCell" Or IsTaggable(node) Then Set c = InnerCell(node)
                If Not c Is Nothing Then
                    If Attr(c, "vertex") = "1" Then
                        If HarvestShape(dom, node, c, f, pageName, commit, batch, src, rows, touched, info, nameIdx, _
                                        plan, cellOf, typeOf, cnt) Then dirty = True
                    End If
                End If
            Next nv
            ' ---- connectors (ConvertToObject changes the tree, so walk a snapshot)
            For Each nv In Children(root)
                Set node = nv
                Set c = Nothing
                If node.nodeName = "mxCell" Or IsTaggable(node) Then Set c = InnerCell(node)
                If Not c Is Nothing Then
                    If Attr(c, "edge") = "1" Then
                        If HarvestEdge(dom, node, c, f, pageName, commit, batch, src, rows, touched, info, rels, byKey, _
                                       cellOf, typeOf, cnt) Then dirty = True
                    End If
                End If
            Next nv
            If commit Then
                Set ids = modCore.NewDict()
                Dim k As Variant
                For Each k In cellOf.Keys
                    If cellOf(k) <> "" And Left$(cellOf(k), 5) <> "(new)" Then ids(cellOf(k)) = True
                Next k
                RegisterView path, pageName, ids, CountRels(PageRoot(pg)), MasterInfo(), "Harvested"
            End If
        End If
    Next pg
    If commit And dirty Then
        On Error Resume Next
        SaveDom dom, path
        If Err.Number <> 0 Then
            Row_ rows, f, "", "", "File", "Error", "", "", "", "", "", "", "", "Error", _
                 "The repository was updated but the new IDs could not be written into the file (" & Err.Description & _
                 "). Harvesting it again links the shapes by name."
            cnt(5) = cnt(5) + 1
        End If
        On Error GoTo 0
    End If
End Sub

' One shape. Returns True when the file needs saving (commit only).
Private Function HarvestShape(dom As Object, node As Object, c As Object, ByVal f As String, ByVal pg As String, _
                              ByVal commit As Boolean, ByVal batch As String, ByVal src As String, rows As Collection, _
                              touched As Object, info As Object, nameIdx As Object, plan As Object, cellOf As Object, _
                              typeOf As Object, cnt() As Long) As Boolean
    Dim cid As String, id As String, tc As String, lbl As String, eaName As String, rec As Variant, k As String
    Dim vals As Object, issues As Collection, it As Variant, pp() As String, newID As String, a As Variant
    Dim blocked As Boolean, cur As Object, r As Long, merged As Object, matches As Object, disp As String
    cid = Attr(node, "id")
    If InStr(Attr(c, "style"), "ea_band=1") > 0 Then Exit Function
    lbl = CleanLabel(CellLabel(node))
    If Not IsTaggable(node) Then
        If lbl <> "" Then Row_ rows, f, pg, cid, "Shape", "Ignored", "", "", lbl, "", "", "", "", "", _
                               "Not a repository shape (no ea_type) - a note or decoration"
        Exit Function
    End If
    id = Attr(node, "ea_id"): tc = Attr(node, "ea_type"): eaName = Attr(node, "ea_name")
    If id = "" And tc = "" Then
        If lbl <> "" Then Row_ rows, f, pg, cid, "Shape", "Ignored", "", "", lbl, "", "", "", "", "", _
                               "Not a repository shape (no ea_type) - a note or decoration"
        Exit Function
    End If
    If Attr(node, "ea_model") <> "" And Attr(node, "ea_model") <> CStr(modCore.Setting("ModelID")) Then
        Row_ rows, f, pg, cid, "Shape", "Error", id, "", lbl, tc, "", "", "", "Error", _
             "Shape was copied from another model's library; use this model's libraries"
        cnt(5) = cnt(5) + 1
        Exit Function
    End If

    ' ---- shape already linked to an entity
    If id <> "" Then
        If Not info.Exists(id) Then
            Row_ rows, f, pg, cid, "Shape", "Info", id, Attr(node, "ea_display"), lbl, tc, "", "", "", "Warning", _
                 "Entity is no longer in the repository; Update diagrams outlines it in red"
            cellOf(cid) = ""
            Exit Function
        End If
        rec = info(id)
        cellOf(cid) = id: typeOf(cid) = rec(0)
        If rec(3) = "Retired" Then
            Row_ rows, f, pg, cid, "Shape", "Info", id, rec(2), rec(1), rec(0), "", "", "", "Warning", "Entity is retired"
        End If
        If lbl = "" Or lbl = rec(1) Then
            cnt(6) = cnt(6) + 1
            Exit Function
        End If
        If lbl = eaName Then
            Row_ rows, f, pg, cid, "Shape", "Info", id, rec(2), rec(1), rec(0), "Name", lbl, rec(1), "", _
                 "Renamed in the repository; Update diagrams refreshes the label"
            cnt(6) = cnt(6) + 1
            Exit Function
        End If
        ' renamed on the diagram
        r = modData.FindEntityRow(id)
        Set cur = modData.ReadEntityRow(r)
        Set merged = modCore.NewDict()
        For Each it In cur.Keys: merged(it) = cur(it): Next it
        merged("Name") = lbl
        Set issues = modRules.ValidateRow(rec(0), merged, nameIdx, id)
        For Each it In issues
            pp = Split(it, "|", 4)
            If pp(0) = "Error" And pp(2) = "Name" Then
                Row_ rows, f, pg, cid, "Shape", "Error", id, rec(2), rec(1), rec(0), "Name", rec(1), lbl, "Error", pp(3)
                cnt(5) = cnt(5) + 1
                Exit Function
            End If
        Next it
        If commit Then
            Set vals = modCore.NewDict()
            vals("Name") = lbl
            Set issues = New Collection
            modData.UpdateEntity id, vals, src, batch, issues, blocked
            If blocked Then
                Row_ rows, f, pg, cid, "Shape", "Error", id, rec(2), rec(1), rec(0), "Name", rec(1), lbl, "Error", _
                     Replace(modRules.IssueText(issues, 2), vbLf, "; ")
                cnt(5) = cnt(5) + 1
                Exit Function
            End If
            IndexRemove nameIdx, rec(0), CStr(rec(1)), id
            IndexAdd nameIdx, rec(0), lbl, id
            info(id) = Array(rec(0), lbl, rec(2), rec(3), modData.MasterRowHash(modData.FindEntityRow(id)))
            node.setAttribute "label", HtmlEsc(lbl)
            node.setAttribute "ea_name", lbl
            touched(rec(0)) = True
            HarvestShape = True
        End If
        Row_ rows, f, pg, cid, "Shape", IIf(commit, "Renamed", "Rename"), id, rec(2), rec(1), rec(0), "Name", rec(1), lbl, "", ""
        cnt(2) = cnt(2) + 1
        Exit Function
    End If

    ' ---- new shape from the stencil (ea_type, no ea_id)
    If Not modMeta.TypeExists(tc) Then
        Row_ rows, f, pg, cid, "Shape", "Error", "", "", lbl, tc, "", "", "", "Error", "Unknown entity type " & tc
        cnt(5) = cnt(5) + 1
        Exit Function
    End If
    If lbl = "" Or lbl = modMeta.TypeName_(tc) Then
        Row_ rows, f, pg, cid, "Shape", "Error", "", "", lbl, tc, "Name", "", "", "Error", _
             "Give the shape a name (double-click it in draw.io)"
        cnt(5) = cnt(5) + 1
        Exit Function
    End If
    typeOf(cid) = tc
    k = tc & "|" & LCase$(lbl)
    ' an existing entity of the same type and name: link the shape to it
    Set matches = modCore.NewDict()
    If nameIdx.Exists(k) Then
        For Each it In nameIdx(k).Keys
            If info.Exists(it) Then
                If info(it)(3) <> "Retired" Then matches(it) = True
            End If
        Next it
    End If
    If matches.Count > 0 Then
        If matches.Count > 1 Then
            Row_ rows, f, pg, cid, "Shape", "Error", "", "", lbl, tc, "", "", "", "Error", _
                 matches.Count & " entities of this type have this name; drag the right one from the Model Library instead"
            cnt(5) = cnt(5) + 1
            cellOf(cid) = ""
            Exit Function
        End If
        id = FirstKey(matches)
        cellOf(cid) = id
        If commit Then
            WriteEntityTags node, id, info
            HarvestShape = True
        End If
        Row_ rows, f, pg, cid, "Shape", IIf(commit, "Linked", "Link"), id, info(id)(2), lbl, tc, "", "", "", "", _
             "Same name and type as an existing entity - the shape is linked to it"
        cnt(3) = cnt(3) + 1
        Exit Function
    End If
    ' the same new name earlier in this harvest
    If plan.Exists(k) Then
        cellOf(cid) = plan(k)
        If commit And Left$(plan(k), 5) <> "(new)" Then WriteEntityTags node, CStr(plan(k)), info: HarvestShape = True
        Row_ rows, f, pg, cid, "Shape", IIf(commit, "Linked", "Link"), IIf(commit, plan(k), ""), "", lbl, tc, "", "", "", "", _
             "Drawn more than once - one entity is created"
        cnt(3) = cnt(3) + 1
        Exit Function
    End If
    ' new entity: Name plus any properties named after the type's attributes (Edit Data in draw.io)
    Set vals = modCore.NewDict()
    vals("Name") = lbl
    For Each a In Array("Description", "Status", "RepositoryArea")
        If Attr(node, CStr(a)) <> "" Then vals(a) = Attr(node, CStr(a))
    Next a
    For Each a In modMeta.AttrsForType(tc)
        If Attr(node, CStr(a)) <> "" Then vals(a) = modData.ToCellValue(CStr(a), Attr(node, CStr(a)))
    Next a
    If Not vals.Exists("Status") Then vals("Status") = "Active"
    If Not vals.Exists("RepositoryArea") Then vals("RepositoryArea") = "Architecture Landscape"
    Set issues = modRules.ValidateRow(tc, vals, nameIdx, "")
    If modRules.HasErrors(issues) Then
        For Each it In issues
            pp = Split(it, "|", 4)
            If pp(0) = "Error" Then
                Row_ rows, f, pg, cid, "Shape", "Error", "", "", lbl, tc, modMeta.AttrName(pp(2)), "", _
                     modCore.Nz(IIf(vals.Exists(pp(2)), vals(pp(2)), ""), ""), "Error", _
                     pp(3) & IIf(modMeta.IsMapped(tc, pp(2)), " - add it with Edit Data (Ctrl+M) on the shape", "")
            End If
        Next it
        cnt(5) = cnt(5) + 1
        cellOf(cid) = ""
        Exit Function
    End If
    If commit Then
        Set issues = New Collection
        newID = modData.CreateEntity(tc, vals, src, batch, issues)
        If newID = "" Then
            Row_ rows, f, pg, cid, "Shape", "Error", "", "", lbl, tc, "", "", "", "Error", _
                 Replace(modRules.IssueText(issues, 2), vbLf, "; ")
            cnt(5) = cnt(5) + 1
            cellOf(cid) = ""
            Exit Function
        End If
        r = modData.FindEntityRow(newID)
        disp = CStr(modData.MasterSheet().Cells(r, modData.MasterCols()("DisplayID")).Value)
        info(newID) = Array(tc, lbl, disp, "Active", modData.MasterRowHash(r))
        IndexAdd nameIdx, tc, lbl, newID
        plan(k) = newID
        cellOf(cid) = newID
        WriteEntityTags node, newID, info
        touched(tc) = True
        HarvestShape = True
        Row_ rows, f, pg, cid, "Shape", "Added", newID, disp, lbl, tc, "(new entity)", "", lbl, "", ""
    Else
        plan(k) = "(new)|" & k
        cellOf(cid) = plan(k)
        Row_ rows, f, pg, cid, "Shape", "New", "", "(on commit)", lbl, tc, "(new entity)", "", lbl, "", ""
    End If
    For Each it In issues
        pp = Split(it, "|", 4)
        If pp(0) = "Warning" Then Row_ rows, f, pg, cid, "Shape", "Info", "", "", lbl, tc, modMeta.AttrName(pp(2)), "", "", _
                                       "Warning", pp(3)
    Next it
    cnt(1) = cnt(1) + 1
End Function

Private Sub WriteEntityTags(node As Object, ByVal id As String, info As Object)
    node.setAttribute "ea_id", id
    node.setAttribute "ea_type", info(id)(0)
    node.setAttribute "ea_display", info(id)(2)
    node.setAttribute "ea_name", info(id)(1)
    node.setAttribute "ea_model", CStr(modCore.Setting("ModelID"))
End Sub

' One connector. Returns True when the file needs saving (commit only).
Private Function HarvestEdge(dom As Object, node As Object, c As Object, ByVal f As String, ByVal pg As String, _
                             ByVal commit As Boolean, ByVal batch As String, ByVal src As String, rows As Collection, _
                             touched As Object, info As Object, rels As Object, byKey As Object, cellOf As Object, _
                             typeOf As Object, cnt() As Long) As Boolean
    Dim cid As String, relID As String, rc As String, sCell As String, tCell As String, sid As String, tid As String
    Dim sType As String, tType As String, desc As String, inferred As Boolean, issues As Collection, newID As String
    Dim lbl As String, o As Object, key As String
    cid = Attr(node, "id")
    If IsTaggable(node) Then
        relID = Attr(node, "ea_rel_id"): rc = Attr(node, "ea_reltype")
    End If
    sCell = Attr(c, "source"): tCell = Attr(c, "target")
    If cellOf.Exists(sCell) Then sid = cellOf(sCell)
    If cellOf.Exists(tCell) Then tid = cellOf(tCell)
    If typeOf.Exists(sCell) Then sType = typeOf(sCell)
    If typeOf.Exists(tCell) Then tType = typeOf(tCell)
    lbl = CleanLabel(CellLabel(node))

    If relID = "" And rc = "" Then
        ' untagged connector: only of interest between two repository shapes
        If sid = "" Or tid = "" Then Exit Function
        rc = RelFromStyle(Attr(c, "style"))
        If rc = "" Then
            Row_ rows, f, pg, cid, "Connector", "Ignored", "", "", ShapeName(sid, info) & " - " & ShapeName(tid, info), "", _
                 "", "", "", "Warning", "Plain connector between repository shapes - use a connector from the ArchiMate 4 library"
            Exit Function
        End If
        inferred = True
    End If
    desc = ShapeName(sid, info) & " " & modMeta.RelVerb(rc) & " " & ShapeName(tid, info)

    ' ---- connector already linked to a relationship
    If relID <> "" Then
        If Not rels.Exists(relID) Then
            Row_ rows, f, pg, cid, "Connector", "Info", relID, "", desc, rc, "", "", "", "Warning", _
                 "Relationship is no longer in the repository"
        ElseIf (sid <> "" And sid <> rels(relID)(1)) Or (tid <> "" And tid <> rels(relID)(2)) Then
            Row_ rows, f, pg, cid, "Connector", "Info", relID, "", desc, rc, "", "", "", "Warning", _
                 "Connector was moved to other shapes; the relationship was not changed. Delete it and draw a new one."
        Else
            cnt(6) = cnt(6) + 1
        End If
        Exit Function
    End If

    ' ---- new relationship
    If sid = "" Or tid = "" Then
        Row_ rows, f, pg, cid, "Connector", "Error", "", "", desc, rc, "", "", "", "Error", _
             "Connector is not attached to two repository shapes at both ends"
        cnt(5) = cnt(5) + 1
        Exit Function
    End If
    If Not modMeta.gRelTypes.Exists(rc) Then
        Row_ rows, f, pg, cid, "Connector", "Error", "", "", desc, rc, "", "", "", "Error", "Unknown relationship type " & rc
        cnt(5) = cnt(5) + 1
        Exit Function
    End If
    If Not modMeta.IsRelAllowed(sType, rc, tType) Then
        Row_ rows, f, pg, cid, "Connector", "Error", "", "", desc, rc, "", "", "", "Error", _
             modMeta.TypeName_(sType) & " " & modMeta.RelVerb(rc) & " " & modMeta.TypeName_(tType) & " is not allowed"
        cnt(5) = cnt(5) + 1
        Exit Function
    End If
    key = rc & "|" & sid & "|" & tid
    If byKey.Exists(key) Then
        If byKey(key) = "(planned)" Then
            Row_ rows, f, pg, cid, "Connector", "Link", "", "", desc, rc, "", "", "", "", _
                 "Drawn more than once - one relationship is created"
            cnt(3) = cnt(3) + 1
            Exit Function
        End If
        If commit Then
            Set o = ConvertToObject(dom, node)
            o.setAttribute "ea_rel_id", byKey(key)
            o.setAttribute "ea_reltype", rc
            HarvestEdge = True
        End If
        Row_ rows, f, pg, cid, "Connector", IIf(commit, "Linked", "Link"), CStr(byKey(key)), "", desc, rc, "", "", "", "", _
             "Already in the repository - the connector is linked to it"
        cnt(3) = cnt(3) + 1
        Exit Function
    End If
    If commit Then
        Set issues = New Collection
        newID = modData.CreateRelationship(rc, sid, tid, lbl, "", src, batch, issues)
        If newID = "" Then
            Row_ rows, f, pg, cid, "Connector", "Error", "", "", desc, rc, "", "", "", "Error", _
                 Replace(modRules.IssueText(issues, 2), vbLf, "; ")
            cnt(5) = cnt(5) + 1
            Exit Function
        End If
        rels(newID) = Array(rc, sid, tid, "Active", lbl)
        byKey(rc & "|" & sid & "|" & tid) = newID
        Set o = ConvertToObject(dom, node)
        o.setAttribute "ea_rel_id", newID
        o.setAttribute "ea_reltype", rc
        touched(sType) = True: touched(tType) = True
        HarvestEdge = True
        Row_ rows, f, pg, cid, "Connector", "Added", newID, "", desc, rc, "", "", "", "", _
             IIf(inferred, "Relationship type taken from the line style", "")
    Else
        byKey(rc & "|" & sid & "|" & tid) = "(planned)"
        Row_ rows, f, pg, cid, "Connector", "New", "", "", desc, rc, "", "", "", "", _
             IIf(inferred, "Relationship type taken from the line style", "")
    End If
    cnt(4) = cnt(4) + 1
End Function

Private Function ShapeName(ByVal ref As String, info As Object) As String
    If ref = "" Then
        ShapeName = "(unattached)"
    ElseIf Left$(ref, 6) = "(new)|" Then
        ShapeName = Mid$(ref, InStrRev(ref, "|") + 1)
    ElseIf info.Exists(ref) Then
        ShapeName = info(ref)(1)
    Else
        ShapeName = ref
    End If
End Function

Private Sub WriteReview(ByVal fileList As String, rows As Collection, ByVal committed As Boolean)
    Dim ws As Worksheet, lo As ListObject, x As Variant, core As String
    Set ws = ThisWorkbook.Worksheets(SH_REVIEW)
    Set lo = ReviewLO()
    core = "'" & ThisWorkbook.Name & "'!modDiagram."
    modCore.BeginBusy
    modCore.UnprotectSheet ws
    modCore.DeleteButtons ws
    modCore.ClearTable lo
    For Each x In rows
        modCore.AppendRow lo, x
    Next x
    ws.Range("A3").Value = IIf(committed, "Committed: ", "Preview: ") & rows.Count & " row(s)"
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").Value = "Files"
    ws.Range("B4").Value = IIf(committed, "", fileList)     ' cleared after a commit so it cannot run twice
    ws.Range("B4").Font.Color = RGB(89, 89, 89)
    If Not committed Then modCore.AddButton ws, "Commit harvest", core & "HarvestCommit", ws.Range("H1"), 150
    modCore.AddButton ws, "Diagrams", core & "ShowDiagrams", ws.Range("J1"), 120
    modCore.AddButton ws, "Back to Home", "'" & ThisWorkbook.Name & "'!modForms.GoHome", ws.Range("L1"), 120
    modCore.ProtectSheet ws
    modCore.EndBusy
    ThisWorkbook.Activate
    ws.Activate
End Sub
