Option Compare Database
Option Explicit
Sub CheckUnitFeatureSpaceRelationships()
Dim mydb As DAO.Database, unitrels As DAO.Recordset, unitFeatures As DAO.Recordset, unitSpaces As DAO.Recordset, unitBuildings As DAO.Recordset
Dim featureSpaces As DAO.Recordset, featureBuildings As DAO.Recordset, getFeatures, getSpaces, getBuildings, thisSpace, ishere
Dim sql, writeToTable As DAO.Recordset, counter
sql = "DELETE from LocalCheckUnitFeatureSpaceRels;"
DoCmd.RunSQL sql
counter = 0
Set mydb = CurrentDb
Set unitrels = mydb.OpenRecordset("Exca: Unit Sheet with Relationships", dbOpenSnapshot)
Set writeToTable = mydb.OpenRecordset("LocalCheckUnitFeatureSpaceRels", dbOpenDynaset)
If Not (unitrels.BOF And unitrels.EOF) Then
    unitrels.MoveFirst
    Do Until unitrels.EOF 'Or counter = 1000
        getFeatures = unitrels![Feature]
        getSpaces = unitrels![Space]
        getBuildings = unitrels![Building]
        If Not IsNull(unitrels![Feature]) Or unitrels![Feature] <> "" Then
            sql = "SELECT * FROM [Exca: Units in Features] where [unit] = " & unitrels![Unit Number] & ";"
            Set unitFeatures = mydb.OpenRecordset(sql, dbOpenSnapshot)
            If Not (unitFeatures.BOF And unitFeatures.EOF) Then
                unitFeatures.MoveFirst
                Do Until unitFeatures.EOF
                    Set featureSpaces = mydb.OpenRecordset("SELECT * FROM [Exca: Features in Spaces] where [Feature] = " & unitFeatures![In_feature], dbOpenSnapshot)
                    If Not (featureSpaces.BOF And featureSpaces.EOF) Then
                        featureSpaces.MoveFirst
                        thisSpace = "," & Trim(featureSpaces![In_Space]) & ","
                        If InStr(getSpaces, thisSpace) > 0 Then
                        Else
                            writeToTable.AddNew
                            writeToTable![UnitNumber] = unitrels![Unit Number]
                            writeToTable![AssociatedWithFeature] = unitFeatures![In_feature]
                            writeToTable![Problem] = "unit: " & unitrels![Unit Number] & " is assoc with feature: " & unitFeatures![In_feature] & ", this feature is in turn related to space: " & featureSpaces![In_Space] & " BUT this unit is NOT related to this space, it is related to '" & getSpaces & "'"
                            writeToTable.Update
                        End If
                    End If
                    featureSpaces.Close
                    Set featureSpaces = Nothing
                    unitFeatures.MoveNext
                Loop
            End If
            unitFeatures.Close
            Set unitFeatures = Nothing
        End If
    unitrels.MoveNext
    counter = counter + 1
    Loop
End If
unitrels.Close
Set unitrels = Nothing
mydb.Close
Set mydb = Nothing
MsgBox "done - " & counter & " records checked"
End Sub
Sub CheckUnitFeatureBuildingRelationships()
Dim mydb As DAO.Database, unitrels As DAO.Recordset, unitFeatures As DAO.Recordset, unitSpaces As DAO.Recordset, unitBuildings As DAO.Recordset
Dim featureSpaces As DAO.Recordset, featureBuildings As DAO.Recordset, getFeatures, getSpaces, getBuildings, thisSpace, ishere
Dim sql, writeToTable As DAO.Recordset, counter, thisBuilding
sql = "DELETE from LocalCheckUnitFeatureBuildingRels;"
DoCmd.RunSQL sql
counter = 0
Set mydb = CurrentDb
Set unitrels = mydb.OpenRecordset("Exca: Unit Sheet with Relationships", dbOpenSnapshot)
Set writeToTable = mydb.OpenRecordset("LocalCheckUnitFeatureBuildingRels", dbOpenDynaset)
If Not (unitrels.BOF And unitrels.EOF) Then
    unitrels.MoveFirst
    Do Until unitrels.EOF 'Or counter = 1000
        getFeatures = unitrels![Feature]
        getSpaces = unitrels![Space]
        getBuildings = unitrels![Building]
        If Not IsNull(unitrels![Feature]) Or unitrels![Feature] <> "" Then
            sql = "SELECT * FROM [Exca: Units in Features] where [unit] = " & unitrels![Unit Number] & ";"
            Set unitFeatures = mydb.OpenRecordset(sql, dbOpenSnapshot)
            If Not (unitFeatures.BOF And unitFeatures.EOF) Then
                unitFeatures.MoveFirst
                Do Until unitFeatures.EOF
                    Set featureBuildings = mydb.OpenRecordset("SELECT * FROM [Exca: Features in Buildings] where [Feature] = " & unitFeatures![In_feature], dbOpenSnapshot)
                    If Not (featureBuildings.BOF And featureBuildings.EOF) Then
                        featureBuildings.MoveFirst
                        thisBuilding = "," & Trim(featureBuildings![In_Building]) & ","
                        If InStr(getBuildings, thisBuilding) > 0 Then
                        Else
                            writeToTable.AddNew
                            writeToTable![UnitNumber] = unitrels![Unit Number]
                            writeToTable![AssociatedWithFeature] = unitFeatures![In_feature]
                            writeToTable![Problem] = "unit: " & unitrels![Unit Number] & " is assoc with feature: " & unitFeatures![In_feature] & ", this feature is in turn is related to building: " & featureBuildings![In_Building] & " BUT this unit is NOT related to this building, it is related to '" & getBuildings & "'"
                            writeToTable.Update
                        End If
                    End If
                    featureBuildings.Close
                    Set featureBuildings = Nothing
                    unitFeatures.MoveNext
                Loop
            End If
            unitFeatures.Close
            Set unitFeatures = Nothing
        End If
    unitrels.MoveNext
    counter = counter + 1
    Loop
End If
unitrels.Close
Set unitrels = Nothing
mydb.Close
Set mydb = Nothing
MsgBox "done - " & counter & " records checked"
End Sub
Sub CheckUnitSpaceBuildingRelationships()
Dim mydb As DAO.Database, unitrels As DAO.Recordset, unitFeatures As DAO.Recordset, unitSpaces As DAO.Recordset, unitBuildings As DAO.Recordset
Dim featureSpaces As DAO.Recordset, featureBuildings As DAO.Recordset, getFeatures, getSpaces, getBuildings, thisSpace, ishere
Dim sql, writeToTable As DAO.Recordset, counter, thisBuilding, spaceBuildings As DAO.Recordset
sql = "DELETE from LocalCheckUnitSpaceBuildingRels;"
DoCmd.RunSQL sql
counter = 0
Set mydb = CurrentDb
Set unitrels = mydb.OpenRecordset("Exca: Unit Sheet with Relationships", dbOpenSnapshot)
Set writeToTable = mydb.OpenRecordset("LocalCheckUnitSpaceBuildingRels", dbOpenDynaset)
If Not (unitrels.BOF And unitrels.EOF) Then
    unitrels.MoveFirst
    Do Until unitrels.EOF 'Or counter = 1000
        getFeatures = unitrels![Feature]
        getSpaces = unitrels![Space]
        getBuildings = unitrels![Building]
        If Not IsNull(unitrels![Space]) Or unitrels![Space] <> "" Then
            sql = "SELECT * FROM [Exca: Units in Spaces] where [unit] = " & unitrels![Unit Number] & ";"
            Set unitSpaces = mydb.OpenRecordset(sql, dbOpenSnapshot)
            If Not (unitSpaces.BOF And unitSpaces.EOF) Then
                unitSpaces.MoveFirst
                Do Until unitSpaces.EOF
                    Set spaceBuildings = mydb.OpenRecordset("SELECT * FROM [Exca: Space Sheet] where [Space Number] = " & unitSpaces![In_Space], dbOpenSnapshot)
                    If Not (spaceBuildings.BOF And spaceBuildings.EOF) Then
                        spaceBuildings.MoveFirst
                        If IsNull(spaceBuildings![Building]) Then
                            thisBuilding = Null
                        Else
                            thisBuilding = "," & Trim(spaceBuildings![Building]) & ","
                        End If
                        If Not IsNull(thisBuilding) Or Not IsNull(getBuildings) Then
                            If InStr(getBuildings, thisBuilding) > 0 Then
                            Else
                                writeToTable.AddNew
                                writeToTable![UnitNumber] = unitrels![Unit Number]
                                writeToTable![AssociatedWithSpace] = unitSpaces![In_Space]
                                writeToTable![Problem] = "unit: " & unitrels![Unit Number] & " is assoc with space: " & unitSpaces![In_Space] & ", this space is in turn related to building: " & spaceBuildings![Building] & " BUT this unit is NOT related to this building, it is related to '" & getBuildings & "'"
                                writeToTable.Update
                            End If
                        End If
                    End If
                    spaceBuildings.Close
                    Set spaceBuildings = Nothing
                    unitSpaces.MoveNext
                Loop
            End If
            unitSpaces.Close
            Set unitSpaces = Nothing
        End If
    unitrels.MoveNext
    counter = counter + 1
    Loop
End If
unitrels.Close
Set unitrels = Nothing
mydb.Close
Set mydb = Nothing
MsgBox "done - " & counter & " records checked"
End Sub
Sub CheckFeatureSpaceBuildingRelationships()
Dim mydb As DAO.Database, featurerels As DAO.Recordset, unitFeatures As DAO.Recordset, unitSpaces As DAO.Recordset, unitBuildings As DAO.Recordset
Dim featureSpaces As DAO.Recordset, featureBuildings As DAO.Recordset, getFeatures, getSpaces, getBuildings, thisSpace, ishere
Dim sql, writeToTable As DAO.Recordset, counter, thisBuilding, spaceBuildings As DAO.Recordset
sql = "DELETE from LocalCheckFeatureSpaceBuildingRels;"
DoCmd.RunSQL sql
counter = 0
Set mydb = CurrentDb
Set featurerels = mydb.OpenRecordset("Exca: Features with Relationships", dbOpenSnapshot)
Set writeToTable = mydb.OpenRecordset("LocalCheckFeatureSpaceBuildingRels", dbOpenDynaset)
If Not (featurerels.BOF And featurerels.EOF) Then
    featurerels.MoveFirst
    Do Until featurerels.EOF 'Or counter = 1000
        If featurerels![Feature Number] = 5000 Then
        End If
        getSpaces = featurerels![Space]
        getBuildings = featurerels![Building]
        If Not IsNull(featurerels![Space]) Or featurerels![Space] <> "" Then
            sql = "SELECT * FROM [Exca: Features in Spaces] where [feature] = " & featurerels![Feature Number] & ";"
            Set featureSpaces = mydb.OpenRecordset(sql, dbOpenSnapshot)
            If Not (featureSpaces.BOF And featureSpaces.EOF) Then
                featureSpaces.MoveFirst
                Do Until featureSpaces.EOF
                    Set spaceBuildings = mydb.OpenRecordset("SELECT * FROM [Exca: Space Sheet] where [Space Number] = " & featureSpaces![In_Space], dbOpenSnapshot)
                    If Not (spaceBuildings.BOF And spaceBuildings.EOF) Then
                        spaceBuildings.MoveFirst
                        If Not IsNull(spaceBuildings![Building]) Then
                            thisBuilding = "," & Trim(spaceBuildings![Building]) & ","
                        Else
                            thisBuilding = Null
                        End If
                        If InStr(getBuildings, thisBuilding) > 0 Then
                        ElseIf Not IsNull(InStr(getBuildings, thisBuilding)) Then
                            writeToTable.AddNew
                            writeToTable![FeatureNumber] = featurerels![Feature Number]
                            writeToTable![AssociatedWithSpace] = featureSpaces![In_Space]
                            writeToTable![Problem] = "feature: " & featurerels![Feature Number] & " is assoc with space: " & featureSpaces![In_Space] & ", this space is in turn related to building: '" & spaceBuildings![Building] & "' BUT this feature NOT related to this building, it is related to '" & getBuildings & "'"
                            writeToTable.Update
                        End If
                    End If
                    spaceBuildings.Close
                    Set spaceBuildings = Nothing
                    featureSpaces.MoveNext
                Loop
            End If
            featureSpaces.Close
            Set featureSpaces = Nothing
        End If
    featurerels.MoveNext
    counter = counter + 1
    Loop
End If
featurerels.Close
Set featurerels = Nothing
mydb.Close
Set mydb = Nothing
MsgBox "done - " & counter & " records checked"
End Sub
Sub CheckFeatureSpaceUnitSpaceRelationships()
Dim mydb As DAO.Database, featurerels As DAO.Recordset, unitFeatures As DAO.Recordset, unitSpaces As DAO.Recordset, unitBuildings As DAO.Recordset
Dim featureSpaces As DAO.Recordset, featureBuildings As DAO.Recordset, getFeatures, getSpaces, getBuildings, thisSpace, ishere
Dim sql, writeToTable As DAO.Recordset, counter, checkpresent, strtoprint
sql = "DELETE from LocalCheckFeatureSpaceUnitSpaceRels;"
DoCmd.RunSQL sql
counter = 0
Dim sqlFeature, response
response = MsgBox("4040 and South only?", vbQuestion + vbYesNo, "Area Filter")
If response = vbYes Then
    sqlFeature = "SELECT * FROM [Exca: Features with Relationships] WHERE Area = 'South' or Area = '4040'"
Else
    sqlFeature = "Exca: Features with Relationships"
End If
Set mydb = CurrentDb
Set featurerels = mydb.OpenRecordset(sqlFeature, dbOpenSnapshot)
Set writeToTable = mydb.OpenRecordset("LocalCheckFeatureSpaceUnitSpaceRels", dbOpenDynaset)
If Not (featurerels.BOF And featurerels.EOF) Then
    featurerels.MoveFirst
    Do Until featurerels.EOF 'Or counter = 1000
        getSpaces = featurerels![Space]
        If Not IsNull(featurerels![Space]) Or featurerels![Space] <> "" Then
            sql = "SELECT * FROM [Exca: Units in Features] where [In_Feature] = " & featurerels![Feature Number] & ";"
            Set unitFeatures = mydb.OpenRecordset(sql, dbOpenSnapshot)
            If Not (unitFeatures.BOF And unitFeatures.EOF) Then
                unitFeatures.MoveFirst
                Do Until unitFeatures.EOF
                    Set unitSpaces = mydb.OpenRecordset("SELECT * FROM [Exca: Units in Spaces] where [Unit] = " & unitFeatures![Unit], dbOpenSnapshot)
                    If Not (unitSpaces.BOF And unitSpaces.EOF) Then
                        unitSpaces.MoveFirst
                        Do Until unitSpaces.EOF
                            thisSpace = "," & Trim(unitSpaces![In_Space]) & ","
                            strtoprint = strtoprint & "," & unitSpaces![In_Space]
                            If InStr(getSpaces, thisSpace) > 0 Then
                            checkpresent = True
                            Exit Do
                            Else
                                checkpresent = False
                            End If
                        unitSpaces.MoveNext
                        Loop
                        If checkpresent = False Then
                            writeToTable.AddNew
                            writeToTable![UnitNumber] = unitFeatures![Unit]
                            writeToTable![AssociatedWithFeature] = unitFeatures![In_feature]
                            writeToTable![Problem] = "Feature: " & featurerels![Feature Number] & " is assoc with spaces: " & getSpaces & ", unit: " & unitFeatures![Unit] & " is associated with this feature but is not associated with any of these spaces, instead it is associated with space/s " & strtoprint
                            writeToTable.Update
                        End If
                        strtoprint = ""
                    End If
                    unitSpaces.Close
                    Set unitSpaces = Nothing
                    unitFeatures.MoveNext
                Loop
            End If
            unitFeatures.Close
            Set unitFeatures = Nothing
        End If
    featurerels.MoveNext
    counter = counter + 1
    Loop
End If
featurerels.Close
Set featurerels = Nothing
mydb.Close
Set mydb = Nothing
MsgBox "done - " & counter & " records checked"
End Sub
