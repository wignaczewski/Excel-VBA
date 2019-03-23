'Moje narzędzia
'
Sub KonwertujKomorkiNaLiczby()
'
'Konwertuje wszystkie wartości w zaznaczonym regionie na String (liczby)
'
    Dim Zakres As Range, Komorka As Range
    Set Zakres = ActiveCell.CurrentRegion
   
    For Each Komorka In Zakres
        If IsNumeric(Komorka) Then
        'Konwertuje komórke na String
        Komorka = CSng(Komorka)

        End If
    Next
End Sub

Sub UsunWzystkieArkusze()
'
'Usuwa wszystkie Arkusze o domyślnej nazwie Arkusz lub Sheet
'
    Dim Arkusz As Object
    Application.DisplayAlerts = False
    For Each Arkusz In Sheets
        If UCase(Left(Arkusz.Name, 6)) = "ARKUSZ" Or UCase(Left(Arkusz.Name, 5)) = "SHEET" Then
            Arkusz.Delete

        End If


        'Arkusz.Protect "haslo"
    Next
    Application.DisplayAlerts = True
End Sub
Sub UkryjArkusze()
'
'Ukrywa akrusze DATA
'
    Dim Arkusz As Object
    Application.DisplayAlerts = False
    For Each Arkusz In Sheets
    Arkusz.Activate
        Arkusz.Range("A1").Select
        If UCase(Left(Arkusz.Name, 4)) = "DATA" Or UCase(Left(Arkusz.Name, 4)) = "DANE" Then
            Arkusz.Visible = False

        End If

    Next

    Worksheets(1).Select
    Application.DisplayAlerts = True
End Sub

Sub OdkryjArkusze()
'
'Odkrywa wszystkie ukryte arkusze
'
    Dim Arkusz As Object
    Application.DisplayAlerts = False
    For Each Arkusz In Sheets
        Arkusz.Visible = True
            Arkusz.Activate
    Arkusz.Range("A1").Select
    Next
    Worksheets(1).Select
    Application.DisplayAlerts = True
End Sub


 ' Module to remove all hidden names on active workbook
   Sub Remove_Hidden_Names()

       ' Dimension variables.
       Dim xName As Variant
       Dim Result As Variant
       Dim Vis As Variant

       ' Loop once for each name in the workbook.
       For Each xName In ActiveWorkbook.Names

           'If a name is not visible (it is hidden)...
           If xName.Visible = True Then
               Vis = "Visible"
           Else
               Vis = "Hidden"
           End If

           ' ...ask whether or not to delete the name.
           Result = MsgBox(prompt:="Delete " & Vis & " Name " & _
               Chr(10) & xName.Name & "?" & Chr(10) & _
               "Which refers to: " & Chr(10) & xName.RefersTo, _
               Buttons:=vbYesNo)

           ' If the result is true, then delete the name.
           If Result = vbYes Then xName.Delete

           ' Loop to the next name.
       Next xName

   End Sub
        
Sub Lokalizacja()
    Debug.Print ThisWorkbook.Path
    Debug.Print Application.StartupPath
End Sub
