Attribute VB_Name = "Module3"
Option Explicit

Sub Main()
'--------------------------------------------------------------------------'
'OPEN ALL THE STP FILES IN A DESIGNATED FOLDER USING LOOP
'--------------------------------------------------------------------------'
Dim strFile As String
Dim strDir As String

strDir = "D:\wang0\Desktop\Dataset\" 'Directory
strFile = Dir(strDir & "*.stp") 'File

Do While strFile <> ""

ThisApplication.Documents.Open (strDir & strFile) 'Open STP file

'--------------------------------------------------------------------------'
'MOVE THE BODY TO RIGHT LOCATION IN RIGHT DIRECTION
'--------------------------------------------------------------------------'
'Get the active part document
Dim mpartDoc As PartDocument
Set mpartDoc = ThisApplication.ActiveDocument

Dim mpartDef As PartComponentDefinition
Set mpartDef = mpartDoc.ComponentDefinition

'Have the user select a body.
Dim mbody As SurfaceBody
Set mbody = mpartDoc.ComponentDefinition.SurfaceBodies.Item(1)

'Extract the bounding box of the imported part
Dim mBox As Box
Set mBox = mbody.RangeBox

'Size of the part
Dim sx As Double
Dim sy As Double
Dim sz As Double
sx = mBox.MaxPoint.x - mBox.MinPoint.x
sy = mBox.MaxPoint.y - mBox.MinPoint.y
sz = mBox.MaxPoint.z - mBox.MinPoint.z

If Not mbody Is Nothing Then
'Create a collection containing the body to move.
Dim bodyCollection As ObjectCollection
Set bodyCollection = ThisApplication.TransientObjects.CreateObjectCollection

Call bodyCollection.Add(mbody)

'Create a move definition.
Dim moveDef As MoveDefinition
Set moveDef = mpartDef.Features.MoveFeatures.CreateMoveDefinition(bodyCollection)

'Rotate the body if necessary
'If sx > sy And sx > sz Then
'Call moveDef.AddRotateAboutAxis(mpartDef.WorkAxes.Item(3), True, pi / 2)
'ElseIf sy > sx And sy > sz Then
'Call moveDef.AddRotateAboutAxis(mpartDef.WorkAxes.Item(2), True, pi / 2)
'ElseIf sz > sx And sz >= sy Then
'Call moveDef.AddRotateAboutAxis(mpartDef.WorkAxes.Item(1), True, pi / 2)
'End If
'Call moveDef.AddRotateAboutAxis(mpartDef.WorkAxes.Item(1), True, pi / 2)

'Move the body so that its centroid of the bottom is coincident with the origin
Dim hx As Double
Dim hy As Double
Dim hz As Double
hx = (mBox.MaxPoint.x + mBox.MinPoint.x) / 2
hy = (mBox.MaxPoint.y + mBox.MinPoint.y) / 2 - (mBox.MaxPoint.y - mBox.MinPoint.y) / 2
hz = (mBox.MaxPoint.z + mBox.MinPoint.z) / 2

'Make the translation
Call moveDef.AddFreeDrag(-hx, -hy, -hz)

'Create the move feature.
Dim move As MoveFeature
Set move = mpartDef.Features.MoveFeatures.Add(moveDef)
End If

'--------------------------------------------------------------------------'
'PREPARE THE PARAMETERS FOR SLICING
'--------------------------------------------------------------------------'
'n is the number of layers required
Dim n As Double
n = 10

'Get the active part document.
Dim opartDoc As PartDocument
Set opartDoc = ThisApplication.ActiveDocument

Dim oDef As ComponentDefinition
Set oDef = opartDoc.ComponentDefinition

' Bounding box of the first body
Dim oBox As Box
Set oBox = opartDoc.ComponentDefinition.SurfaceBodies.Item(1).RangeBox

'Create a sketch on the bottom of the part.
Dim oSketch As PlanarSketch
Set oSketch = oDef.Sketches.Add(oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(2), oBox.MinPoint.y))

'Set a reference to the transient geometry object.
Dim oTransGeom As TransientGeometry
Set oTransGeom = ThisApplication.TransientGeometry

'Create a square on the sketch.
Call oSketch.SketchLines.AddAsTwoPointRectangle(oTransGeom.CreatePoint2d((oBox.MaxPoint.x) * 2, (oBox.MaxPoint.z) * 2), oTransGeom.CreatePoint2d(-(oBox.MaxPoint.x) * 2, -(oBox.MaxPoint.z) * 2))

'Create the profile.
Dim oProfile As Profile
Set oProfile = oSketch.Profiles.AddForSolid
Dim i As Integer
Dim x As Double
Dim y As Double
Dim z As Double
Dim w As Double
'x is the total height of the part
'y is the thickness of the cutting plane
'w is the height of the first workpoint
x = oBox.MaxPoint.y - oBox.MinPoint.y
y = x / n
w = y

'Create the cutting plate by extrusion.
Dim oExtrude As ExtrudeFeature
Set oExtrude = oDef.Features.ExtrudeFeatures.AddByDistanceExtent(oProfile, y, kNegativeExtentDirection, kNewBodyOperation)

'Get the start face of the extrude.
Dim oFace As Face
Set oFace = oExtrude.StartFaces(1)

'Get two adjacent edges on the start face.
Dim oEdge1, oEdge2 As Edge
Set oEdge1 = oFace.Edges(1)
Set oEdge2 = oFace.Edges(2)

Dim oTG As TransientGeometry
Set oTG = ThisApplication.TransientGeometry
Dim oWorkPoint As WorkPoint

' Get the slicing locations
For i = 1 To n
Set oWorkPoint = opartDoc.ComponentDefinition.WorkPoints.AddFixed(oTG.CreatePoint(0, w, 0))
w = w + y
Next

'--------------------------------------------------------------------------'
'SLICING OPERATION (BOOLEAN OPERATION)
'--------------------------------------------------------------------------'
'Get the active part document
Dim partDoc As PartDocument
Set partDoc = ThisApplication.ActiveDocument
Dim partDef As PartComponentDefinition
Set partDef = partDoc.ComponentDefinition

Dim tg As TransientGeometry
Set tg = ThisApplication.TransientGeometry
Dim tb As TransientBRep
Set tb = ThisApplication.TransientBRep
Dim tObjs As TransientObjects
Set tObjs = ThisApplication.TransientObjects
 
'Have the bodies selected.
Dim baseBody As SurfaceBody
Set baseBody = partDoc.ComponentDefinition.SurfaceBodies.Item(1)

Dim toolBody As SurfaceBody
Set toolBody = partDoc.ComponentDefinition.SurfaceBodies.Item(2)
                
'Copy the two bodies to create transient copies.
Dim transBase As SurfaceBody
Set transBase = tb.Copy(baseBody)
Dim transTool As SurfaceBody
Set transTool = tb.Copy(toolBody)

'Create a matrix and a point to use in positioning
'the punch.  The matrix is initialized to an identity
'matrix the point is (0,0,0).
Dim trans As Matrix
Set trans = tg.CreateMatrix
Dim lastPosition As Point
Set lastPosition = tg.CreatePoint

'By performing a boolean with the tool at that location.
Dim k As Integer
For k = 2 To partDef.WorkPoints.Count
Dim wp As WorkPoint
Set wp = partDef.WorkPoints.Item(k)
'Transform the tool body to the position of slicing. The
'boolean operation is at the last operation so the transform defines
'the difference between the last and the current.
trans.Cell(1, 4) = wp.Point.x - lastPosition.x
trans.Cell(2, 4) = wp.Point.y - lastPosition.y
Call tb.Transform(transTool, trans)

'Do the boolean operation.
'Call tb.DoBoolean(transBase, transTool, kBooleanTypeDifference)
Call tb.DoBoolean(transBase, transTool, kBooleanTypeIntersect)

'Save the last position.
Set lastPosition = wp.Point

'Create a base body feature of the result.
Dim nonParamFeatures As NonParametricBaseFeatures
Set nonParamFeatures = partDef.Features.NonParametricBaseFeatures
Dim nonParamDef As NonParametricBaseFeatureDefinition
Set nonParamDef = nonParamFeatures.CreateDefinition

'Save the generated slice
Dim objs As ObjectCollection
Set objs = tObjs.CreateObjectCollection
Call objs.Add(transBase)
nonParamDef.BRepEntities = objs
nonParamDef.OutputType = kSolidOutputType
Call nonParamFeatures.AddByDefinition(nonParamDef)

'Reset the basebody
Set transBase = tb.Copy(baseBody)
'Set transTool = tb.Copy(toolBody)
Next

'Turn off the display of the original two
'features to see the result.
baseBody.Visible = False
toolBody.Visible = False

'Make all bodies invisible
Dim obody As SurfaceBody
For Each obody In partDef.SurfaceBodies
obody.Visible = False
Next

'Force a refresh of the view.
ThisApplication.ActiveDocument.Update

'--------------------------------------------------------------------------'
'SAVE THE CROSSSECTION SHAPE OF EACH SLICE AS STL FILE
'--------------------------------------------------------------------------'
'Get the active part document
Dim oApp As Application
Set oApp = ThisApplication
Dim oPart As PartDocument
Set oPart = oApp.ActiveDocument
Dim oCompDef As PartComponentDefinition
Set oCompDef = oPart.ComponentDefinition

'Have each slice body selected by For loop
Dim body As SurfaceBody
'For Each body In oCompDef.SurfaceBodies
Dim j As Integer
For j = 3 To n + 2
Set body = oPart.ComponentDefinition.SurfaceBodies.Item(j)
body.Visible = True

'Bounding box of each slice body
Dim xBox As Box
Set xBox = body.RangeBox
'Set xBox = calculateTightBoundingBox(body)

'Centroid of the bottom of the slice
Dim centroid_x As Double
Dim centroid_y_bottom As Double
Dim centroid_y_top As Double
Dim centroid_z As Double
centroid_x = (xBox.MaxPoint.x + xBox.MinPoint.x) / 2
centroid_y_bottom = xBox.MinPoint.y
centroid_y_top = xBox.MinPoint.y + (n - (j - 2) + 1) * y
centroid_z = (xBox.MaxPoint.z + xBox.MinPoint.z) / 2

If body.Visible Then
body.Name = "Slice" & n - (j - 2) + 1
Call SaveAsTXT(body, centroid_x, centroid_y_bottom, centroid_y_top, centroid_z)
'Make the solid body invisible so that only the current value is saved as TXT file
body.Visible = False
End If

Next

'Close the completed part file
'ThisApplication.Documents.CloseAll

'--------------------------------------------------------------------------'
'GO TO NEXT ITERATION
'--------------------------------------------------------------------------'
'Go to the next STP part file
strFile = Dir

Loop

End Sub

'--------------------------------------------------------------------------'
'DEFINED FUNCTIONS
'--------------------------------------------------------------------------'
Public Function pi() As Double

pi = 4 * Atn(1)
End Function

'Return the path of the input filename.
Public Function FilePath(ByVal fullFilename As String) As String
    'Extract the path by getting everything up to and
    'including the last backslash "\".
    FilePath = Left$(fullFilename, InStrRev(fullFilename, "\") - 1)
End Function


'Return the name of the file, without the path.
Public Function Filename(ByVal fullFilename As String) As String
    ' Extract the filename by getting everything to
    ' the right of the last backslash.
    Filename = Right$(fullFilename, Len(fullFilename) - _
               InStrRev(fullFilename, "\"))
End Function


'Return the base name of the input filename, without
'the path or the extension.
Public Function BaseFilename(ByVal fullFilename As String) As String
    'Extract the filename by getting everttgubg to
    'the right of the last backslash.
    Dim temp As String
    temp = Right$(fullFilename, Len(fullFilename) - _
           InStrRev(fullFilename, "\"))

    'Get the base filename by getting everything to
    'the left of the last period ".".
    BaseFilename = Left$(temp, InStrRev(temp, ".") - 1)
End Function


'Return the extension of the input filename.
Public Function FileExtension(ByVal fullFilename As String) As String
    'Extract the filename by getting everthing to
    'the right of the last backslash.
    Dim temp As String
    temp = Right$(fullFilename, Len(fullFilename) - _
           InStrRev(fullFilename, "\"))

    'Get the base filename by getting everything to
    'the right of the last period ".".
    FileExtension = Right$(temp, Len(temp) - InStrRev(temp, ".") + 1)
End Function


'Function of saving as TXT files
Sub SaveAsTXT(body As SurfaceBody, centroid_x As Double, centroid_y_bottom As Double, centroid_y_top As Double, centroid_z As Double)
Const SEEDFilePath As String = "D:\wang0\Desktop\SEED\"
Dim partDoc As Document
Set partDoc = ThisApplication.ActiveDocument
Dim fFileName As String
fFileName = partDoc.fullFilename
Dim Filename As String
Filename = SEEDFilePath & BaseFilename(fFileName) & "_" & body.Name & ".txt"
Dim My_filenumber As Integer
My_filenumber = FreeFile
Open Filename For Output As #My_filenumber
Write #My_filenumber, centroid_x, centroid_y_bottom, centroid_y_top, centroid_z
Close #My_filenumber
End Sub























