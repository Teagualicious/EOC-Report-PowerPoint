"""VBA injection for the Search Dashboard.

At export time (Windows + Excel + pywin32), inject_search_vba() opens the
freshly written workbook, adds the VBA search engine, and saves it as
.xlsm. Requires the one-time Excel setting:
    File > Options > Trust Center > Trust Center Settings >
    Macro Settings > "Trust access to the VBA project object model"

On any failure the original .xlsx is left untouched and (path, reason) is
returned so the UI can tell the user exactly what to enable.

VBA behavior (mirrors the grammar documented in excel_search_dashboard):
  - Typing in the search cell + Enter, or clicking the Search button chip,
    parses comma-separated terms against _SearchIndex and rebuilds the
    results table below.
  - Suggestion chips for the last term fragment render in row 7; clicking
    one completes the term and re-runs.
  - Clicking a date 📅 chip opens a small runtime-built calendar UserForm.
  - Additive metrics sum; ratio metrics average (per _Config's agg column).
"""

import logging
import os

logger = logging.getLogger(__name__)


class VBAInjectionError(Exception):
    pass


_VBA_SRC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vba_src")


def _load_vba(filename):
    with open(os.path.join(_VBA_SRC_DIR, filename), "r", encoding="utf-8") as f:
        return f.read()


def VBA_MOD_SEARCH():
    return _load_vba("modSearch.bas")

VBA_SHEET_SEARCH = r'''
Private Sub txtSearch_Change()
    ' Fires on EVERY keystroke — live suggestions while typing
    On Error Resume Next
    modSearch.UpdateSuggestionsFor Me.OLEObjects("txtSearch").Object.Text
End Sub

Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                              ByVal Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then  ' Enter commits -> Worksheet_Change -> RunSearch
        Me.Range("SearchCell").Value = Me.OLEObjects("txtSearch").Object.Text
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Not Intersect(Target, Me.Range("SearchCell")) Is Nothing Then
        modSearch.UpdateSuggestions
        modSearch.RunSearch
    ElseIf Not Intersect(Target, Union(Me.Range("DateStart"), _
                                       Me.Range("DateEnd"))) Is Nothing Then
        modSearch.RunSearch
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    If Target.Cells.Count > 1 Then Exit Sub
    ' Search button chip
    If Target.Address = Me.Range("H6").Address Then
        Dim t As String
        t = ""
        On Error Resume Next
        t = Me.OLEObjects("txtSearch").Object.Text
        On Error GoTo 0
        If t <> "" Then Me.Range("SearchCell").Value = t  ' triggers RunSearch
        If t = "" Then modSearch.RunSearch
        Me.Range("SearchCell").Select
    ' Calendar chips
    ElseIf Target.Address = Me.Range("K6").Address Then
        modSearch.ShowCalendarFor "DateStart"
        Me.Range("DateStart").Select
    ElseIf Target.Address = Me.Range("N6").Address Then
        modSearch.ShowCalendarFor "DateEnd"
        Me.Range("DateEnd").Select
    ' Copy-results chip
    ElseIf Target.Address = Me.Range("D11").Address Then
        modSearch.CopyResults
        Me.Range("SearchCell").Select
    ' Suggestion chips (row 7, cols B..)
    ElseIf Target.Row = 9 And Target.Column >= 2 And Target.Column <= 40 Then
        If CStr(Target.Value) <> "" Then
            modSearch.ApplySuggestion CStr(Target.Value)
            Me.Range("SearchCell").Select
        End If
    End If
End Sub
'''

# UserForm code — the form builds its own controls at runtime so no
# designer-layout manipulation is needed during injection.
VBA_FORM_CALENDAR = r'''
Public TargetRangeName As String
Private mYear As Integer
Private mMonth As Integer

Private Sub UserForm_Initialize()
    Me.Caption = "Pick a date"
    Me.Width = 232: Me.Height = 220
    mYear = Year(Date): mMonth = Month(Date)
    BuildGrid
End Sub

Private Sub BuildGrid()
    ' Clear safely — removing during For Each iteration skips controls
    Do While Me.Controls.Count > 0
        Me.Controls.Remove Me.Controls(0).Name
    Loop

    Dim btnPrev As Object, btnNext As Object, lblTitle As Object
    Set btnPrev = Me.Controls.Add("Forms.CommandButton.1", "btnPrev")
    btnPrev.Caption = "<": btnPrev.Left = 6: btnPrev.Top = 6
    btnPrev.Width = 24: btnPrev.Height = 18
    Set btnNext = Me.Controls.Add("Forms.CommandButton.1", "btnNext")
    btnNext.Caption = ">": btnNext.Left = 192: btnNext.Top = 6
    btnNext.Width = 24: btnNext.Height = 18
    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle")
    lblTitle.Caption = Format$(DateSerial(mYear, mMonth, 1), "mmmm yyyy")
    lblTitle.Left = 36: lblTitle.Top = 8: lblTitle.Width = 150
    lblTitle.TextAlign = 2

    Dim dows As Variant: dows = Array("Su", "Mo", "Tu", "We", "Th", "Fr", "Sa")
    Dim i As Integer
    For i = 0 To 6
        Dim lbl As Object
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblD" & i)
        lbl.Caption = dows(i): lbl.Left = 8 + i * 30: lbl.Top = 30
        lbl.Width = 26: lbl.TextAlign = 2: lbl.Font.Bold = True
    Next i

    Dim firstDow As Integer, daysIn As Integer, d As Integer
    firstDow = Weekday(DateSerial(mYear, mMonth, 1)) - 1
    daysIn = Day(DateSerial(mYear, mMonth + 1, 0))
    For d = 1 To daysIn
        Dim btn As Object, slot As Integer
        slot = firstDow + d - 1
        Set btn = Me.Controls.Add("Forms.CommandButton.1", "day_" & d)
        btn.Caption = CStr(d)
        btn.Left = 8 + (slot Mod 7) * 30
        btn.Top = 46 + (slot \ 7) * 26
        btn.Width = 26: btn.Height = 22
    Next d
End Sub

'''

# Runtime-created buttons cannot have compile-time event handlers; a tiny
# WithEvents class relays clicks from every runtime button (days AND the
# month-navigation arrows — compile-time btnPrev_Click/btnNext_Click subs
# would never fire for controls added via Controls.Add).
VBA_CLASS_DAYBTN = r'''
Public WithEvents Btn As MSForms.CommandButton
Public ParentForm As Object

Private Sub Btn_Click()
    If Btn.Name = "btnPrev" Then
        ParentForm.NavMonth -1
    ElseIf Btn.Name = "btnNext" Then
        ParentForm.NavMonth 1
    Else
        ParentForm.PickDay CInt(Mid$(Btn.Name, 5))
    End If
End Sub
'''

# Declarations must precede ALL procedures in a VBA module — these are
# hoisted to the very top of the composed form source.
VBA_FORM_CALENDAR_DECLS = "Private mHandlers As Collection\n"

# Additions to the form to wire the relay class
VBA_FORM_CALENDAR_WIRING = r'''
Public Sub PickDay(d As Integer)
    On Error Resume Next
    ThisWorkbook.Worksheets("Search").Range(TargetRangeName).Value = _
        Format$(DateSerial(mYear, mMonth, d), "yyyy-mm-dd")
    Unload Me
    modSearch.RunSearch
End Sub

Public Sub NavMonth(delta As Integer)
    mMonth = mMonth + delta
    If mMonth = 0 Then mMonth = 12: mYear = mYear - 1
    If mMonth = 13 Then mMonth = 1: mYear = mYear + 1
    BuildGrid
End Sub

Private Sub WireHandlers()
    Set mHandlers = New Collection
    Dim c As Control
    For Each c In Me.Controls
        If Left$(c.Name, 4) = "day_" Or c.Name = "btnPrev" _
                Or c.Name = "btnNext" Then
            Dim h As clsDayBtn
            Set h = New clsDayBtn
            Set h.Btn = c
            Set h.ParentForm = Me
            mHandlers.Add h
        End If
    Next c
End Sub
'''


def _compose_calendar_form():
    """Merge the calendar form source with the handler wiring.

    VBA requires module-level declarations BEFORE any procedure, so the
    mHandlers declaration is hoisted above the existing declarations by
    prepending it to the whole source (whose own decls already lead).
    WireHandlers is spliced into the end of every grid rebuild."""
    src = VBA_FORM_CALENDAR_DECLS + VBA_FORM_CALENDAR + VBA_FORM_CALENDAR_WIRING
    src = src.replace("    Next d\nEnd Sub",
                      "    Next d\n    WireHandlers\nEnd Sub", 1)
    return src


def inject_search_vba(xlsx_path):
    """Inject the search VBA and save as .xlsm.

    Returns (final_path, error_reason). error_reason is None on success;
    on failure the original xlsx is untouched and reason explains why.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return xlsx_path, ("pywin32 not available — workbook saved without "
                           "the interactive search (data intact).")

    xlsm_path = os.path.splitext(xlsx_path)[0] + ".xlsm"
    app = None
    wb = None
    try:
        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(os.path.abspath(xlsx_path))

        try:
            vbp = wb.VBProject
            _ = vbp.VBComponents.Count  # touch to trigger trust check
        except Exception:
            return xlsx_path, (
                "Excel blocked VBA injection. Enable it once:\n"
                "File > Options > Trust Center > Trust Center Settings >\n"
                "Macro Settings > check 'Trust access to the VBA project "
                "object model'\nthen re-run the export.")

        # Standard module
        mod = vbp.VBComponents.Add(1)  # vbext_ct_StdModule
        mod.Name = "modSearch"
        # strip the Attribute line (only valid in exported files)
        mod.CodeModule.AddFromString(
            VBA_MOD_SEARCH().replace('Attribute VB_Name = "modSearch"\n', ''))

        # Day-button relay class
        cls = vbp.VBComponents.Add(2)  # vbext_ct_ClassModule
        cls.Name = "clsDayBtn"
        cls.CodeModule.AddFromString(VBA_CLASS_DAYBTN)

        # Calendar form (controls are built at runtime by its own code)
        frm = vbp.VBComponents.Add(3)  # vbext_ct_MSForm
        frm.Name = "frmCalendar"
        frm.CodeModule.AddFromString(_compose_calendar_form())

        # Sheet code-behind for "Search"
        for comp in vbp.VBComponents:
            try:
                if comp.Type == 100 and comp.Properties("Name").Value == "Search":
                    comp.CodeModule.AddFromString(VBA_SHEET_SEARCH)
                    break
            except Exception:
                continue

        # ActiveX search box overlay — enables per-keystroke suggestions.
        # Guarded: if control creation fails, the plain cell still works.
        try:
            ws = wb.Worksheets("Search")
            rng = ws.Range("B6:F6")
            ole = ws.OLEObjects().Add(ClassType="Forms.TextBox.1")
            # Named position args are unreliable via pywin32 (box landed at
            # 0,0 over the banner) — set properties after creation instead
            ole.Left = rng.Left + 1
            ole.Top = rng.Top + 1
            ole.Width = rng.Width - 2
            ole.Height = rng.Height - 2
            ole.Placement = 1  # xlMoveAndSize
            ole.Name = "txtSearch"
            ole.Object.Font.Size = 10   # 12 clipped descenders in the box
            try:
                # long searches wrap in the box; Enter still commits
                ole.Object.MultiLine = True
                ole.Object.WordWrap = True
                ole.Object.EnterKeyBehavior = False
            except Exception:
                pass
            try:
                ole.Object.Font.Name = "Segoe UI"
            except Exception:
                pass
            ole.Object.BorderStyle = 0      # frameless — the cell draws the box
            ole.Object.SpecialEffect = 0    # flat
        except Exception:
            logger.warning("Search textbox overlay not created — falling back "
                           "to cell input (suggestions update on Enter)",
                           exc_info=True)

        wb.SaveAs(os.path.abspath(xlsm_path), FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
        wb.Close(SaveChanges=False)
        wb = None
        import time
        for _ in range(6):
            try:
                os.remove(xlsx_path)
                break
            except OSError:
                time.sleep(0.35)
        else:
            logger.warning("Stale .xlsx left beside the .xlsm (locked): %s",
                           xlsx_path)
        return xlsm_path, None

    except Exception as e:
        logger.exception("VBA injection failed")
        return xlsx_path, f"VBA injection failed: {e}"
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
