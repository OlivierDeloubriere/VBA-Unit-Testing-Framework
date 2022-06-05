Attribute VB_Name = "VBAObjectsAPI_HelperFunctions"
'@Folder("VBA_Objects_API")
Option Explicit

Public Function ObjectHasMethod(ByVal obj As Object, methodName As String) As Boolean
    Dim objectMethods As Variant
    Dim methodIndex As Long
    Dim methodFound As Boolean
    
    objectMethods = GetObjectFunctions(obj, VbMethod)
    
    If IsArray(objectMethods) Then
        For methodIndex = LBound(objectMethods, 1) To UBound(objectMethods, 1)
            If objectMethods(methodIndex, 0) = methodName Then
                methodFound = True
                Exit For
            End If
        Next methodIndex
    End If
    ObjectHasMethod = methodFound
End Function


Public Function ObjectHasProperty(ByVal obj As Object, propertyName As String) As Boolean
    Dim objectProperties As Variant
    Dim propertyIndex As Long
    Dim propertyFound As Boolean
    
    objectProperties = GetObjectFunctions(obj, VbGet)
    
    If IsArray(objectProperties) Then
        For propertyIndex = LBound(objectProperties, 1) To UBound(objectProperties, 1)
            If objectProperties(propertyIndex, 0) = propertyName Then
                propertyFound = True
                Exit For
            End If
        Next propertyIndex
    End If
    ObjectHasProperty = propertyFound
End Function
