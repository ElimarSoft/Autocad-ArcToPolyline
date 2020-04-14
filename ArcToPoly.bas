Attribute VB_Name = "ArcToPoly"
Option Explicit
Dim Dwg1 As AcadDocument
Const StepsPerRevolution As Integer = 90
Public Sub ConvertArcToPoly()
Dim SSet As AcadSelectionSet
Dim item As AcadSelectionSet
Dim ent1 As AcadEntity


   Dim SelName As String
   SelName = "Sel1"
    
   Set Dwg1 = ThisDrawing
   
   'Select objects
    If ThisDrawing.SelectionSets.Count > 0 Then
        For Each item In ThisDrawing.SelectionSets
            If item.Name = SelName Then
                item.Delete
                Exit For
            End If
        Next
    End If
    
    Set SSet = ThisDrawing.SelectionSets.Add(SelName)
    SSet.SelectOnScreen
    Debug.Print SSet.Count
    
    'Process selection
    For Each ent1 In SSet
    
        Debug.Print ent1.ObjectName
    
        If ent1.ObjectName = "AcDbArc" Then
             ArcToPoly ent1
        ElseIf ent1.ObjectName = "AcDbPolyline" Then
             ProcessPolyArcs ent1
        End If
    
    Next


End Sub
Private Sub ProcessPolyArcs(ent1 As AcadEntity)

            Dim Pol1 As AcadLWPolyline
            Dim entArray() As AcadEntity
            Dim n As Integer
            Dim PrevLayer As String
            Dim TempLayer As String
            
            TempLayer = "Temp1"
            
            Set Pol1 = ent1
            
            Dwg1.Layers.Add (TempLayer)
            
            PrevLayer = Pol1.Layer
            Pol1.Layer = TempLayer
            
            entArray = Pol1.Explode
            
            Pol1.Delete
            
            For n = 0 To UBound(entArray)
                If entArray(n).ObjectName = "AcDbArc" Then
                   Set entArray(n) = ArcToPoly(entArray(n))
                   entArray(n).Layer = "Temp1"
                End If
            Next

            Dwg1.Regen (acAllViewports)

            Dim FilterType(0) As Integer
            Dim FilterData(0) As Variant
     
            FilterType(0) = 8 'LayerName
            FilterData(0) = TempLayer

            Dim SSet As AcadSelectionSet
            
            Set SSet = Dwg1.SelectionSets.Add("PolySel3")
            Call SSet.Select(acSelectionSetAll, , , FilterType, FilterData)
            SSet.Delete
            
            Call Dwg1.SendCommand("PEDIT M P  Y J 1 X " + vbCr)
            Dwg1.ModelSpace.item(Dwg1.ModelSpace.Count - 1).Layer = PrevLayer
            
            Dwg1.Layers(TempLayer).Delete
            
End Sub

Private Function ArcToPoly(ent1 As AcadEntity) As AcadLWPolyline

        Dim Arc1 As AcadArc
        Dim Line1 As AcadLine
        Dim Line2 As AcadLine
        Dim Poly1 As AcadLWPolyline
        Dim Poly2 As AcadLWPolyline
        Dim n
        
        Set Arc1 = ent1
        
        Dwg1.SendCommand ("Polygon " + CStr(StepsPerRevolution) + " " + CStr(Arc1.Center(0)) + "," + CStr(Arc1.Center(1)) + " I R " + CStr(Arc1.Radius) + vbCr)
        Set Poly1 = Dwg1.ModelSpace.item(Dwg1.ModelSpace.Count - 1)
        
        'Calculate closest points
        Dim Dist1 As Double
        Dim Dist2 As Double
        Dim DistMin1 As Double
        Dim DistMin2 As Double
        Dim PosMin1 As Integer
        Dim PosMin2 As Integer
        Dim StartPoint(2) As Variant
        Dim EndPoint(2) As Variant
        
        If Arc1.Normal(2) > 0 Then
        
            For n = 0 To 2
                StartPoint(n) = Arc1.StartPoint(n)
                EndPoint(n) = Arc1.EndPoint(n)
            Next
            
        
        Else
        
            For n = 0 To 2
                EndPoint(n) = Arc1.StartPoint(n)
                StartPoint(n) = Arc1.EndPoint(n)
            Next
        
        End If
        
        
        
        For n = 0 To UBound(Poly1.Coordinates) Step 2
            Dist1 = (Poly1.Coordinates(n) - StartPoint(0)) ^ 2 + (Poly1.Coordinates(n + 1) - StartPoint(1)) ^ 2
            Dist2 = (Poly1.Coordinates(n) - EndPoint(0)) ^ 2 + (Poly1.Coordinates(n + 1) - EndPoint(1)) ^ 2
        
            If n = 0 Then
                DistMin1 = Dist1
                DistMin2 = Dist2
                PosMin1 = 0
                PosMin2 = 0
            Else
                If Dist1 < DistMin1 Then
                    DistMin1 = Dist1
                    PosMin1 = n
                End If
                
                If Dist2 < DistMin2 Then
                    DistMin2 = Dist2
                    PosMin2 = n
                End If
            
            End If
        
        Next n
        '*******************************************************************************************************
        
        Dim Point1(1) As Double
        Dim NewCoord() As Double
        Dim Ptr1 As Integer
        Dim Ptr2 As Integer
              
        ReDim NewCoord(UBound(Poly1.Coordinates))
        
        Ptr1 = PosMin1
        Ptr2 = 0
       
        Do
        
            NewCoord(Ptr2) = Poly1.Coordinates(Ptr1)
            NewCoord(Ptr2 + 1) = Poly1.Coordinates(Ptr1 + 1)
            
            Ptr1 = Ptr1 + 2
            Ptr2 = Ptr2 + 2
            
            If Ptr1 > UBound(Poly1.Coordinates) Then Ptr1 = 0
            
            If Ptr1 = PosMin2 Then Exit Do
            
        Loop
        
        NewCoord(0) = StartPoint(0)
        NewCoord(1) = StartPoint(1)
                
        NewCoord(Ptr2) = EndPoint(0)
        NewCoord(Ptr2 + 1) = EndPoint(1)
        
        ReDim Preserve NewCoord(Ptr2 + 1)
        
        Poly1.Delete
        Arc1.Delete
        
        Set Poly2 = Dwg1.ModelSpace.AddLightWeightPolyline(NewCoord)

        Set ArcToPoly = Poly2


End Function
