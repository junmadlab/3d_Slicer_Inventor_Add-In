Attribute VB_Name = "Module4"
Option Explicit

Sub Main()
'--------------------------------------------------------------------------'
'OPEN ALL THE STP FILES IN A DESIGNATED FOLDER USING LOOP
'--------------------------------------------------------------------------'
Dim strFile As String
Dim strDir As String

strDir = "D:\wang0\Desktop\Dataset_full\" 'Directory
strFile = Dir(strDir & "*.stp") 'File

Do While strFile <> ""

ThisApplication.Documents.Open (strDir & strFile) 'Open STP file

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
Set oSketch = oDef.Sketches.Add(oDef.WorkPlanes.AddByPlaneAndOffset(oDef.WorkPlanes(2), 0))

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
x = 1
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

'Centroid of the bottom of the slice
Dim centroid_x As Double
Dim centroid_y As Double
Dim centroid_z As Double
centroid_x = 0
centroid_y = y * (j - 3)
centroid_z = 0

Dim xwp As WorkPoint
Set xwp = oCompDef.WorkPoints.AddFixed(ThisApplication.TransientGeometry.CreatePoint(centroid_x, centroid_y, centroid_z))

Dim pt As Point
Set pt = xwp.Point

If body.Visible Then
'Save each slice as STP file if needed
'body.Name = "Slice" & n - (j - 2) + 1
'Call SaveAsSTP(body)
'body.Visible = False
'Find the bottom face of each slice using the centroid point
Dim oFaces As Object
Set oFaces = body.LocateUsingPoint(kFaceObject, pt, y / (2 * n))

'Put a sketch on the bottom face of each slice
Dim xSketch As PlanarSketch
Set xSketch = oCompDef.Sketches.Add(oFaces, True)
'Set xSketch = oCompDef.Sketches.Add(XZPlane, xBox.MinPoint.y)
'Set xSketch = oCompDef.Sketches.Add(XZPlane, True)
'create Boundary Patch Definition
Dim oBoundaryPatchDef As BoundaryPatchDefinition
Set oBoundaryPatchDef = oCompDef.Features.BoundaryPatchFeatures.CreateBoundaryPatchDefinition

Dim xProfile As Profile
Set xProfile = xSketch.Profiles.AddForSolid

Call oBoundaryPatchDef.BoundaryPatchLoops.Add(xProfile)

'Create the boundary patch feature based on the definition.
Dim oBoundaryPatch As BoundaryPatchFeature
Set oBoundaryPatch = oCompDef.Features.BoundaryPatchFeatures.Add(oBoundaryPatchDef)

'Make the solid body invisible so that only the surface body is saved as STL file
body.Visible = False

Dim oSTLTranslator As TranslatorAddIn
Set oSTLTranslator = ThisApplication.ApplicationAddIns.ItemById("{533E9A98-FC3B-11D4-8E7E-0010B541CD80}")

Dim oContext As TranslationContext
Set oContext = ThisApplication.TransientObjects.CreateTranslationContext

Dim oOptions As NameValueMap
Set oOptions = ThisApplication.TransientObjects.CreateNameValueMap

If oSTLTranslator.HasSaveCopyAsOptions(oPart, oContext, oOptions) Then
' Set accuracy.
' 2 = High,  1 = Medium,  0 = Low
oOptions.Value("Resolution") = 1

' Set output file type:
'   0 - binary,  1 - ASCII
oOptions.Value("OutputFileType") = 0

oContext.Type = kFileBrowseIOMechanism

'Save the bottom face shape as STL file
Dim oData As DataMedium
Set oData = ThisApplication.TransientObjects.CreateDataMedium
Const STLFilePath As String = "D:\wang0\Desktop\STL\"
oBoundaryPatch.Name = "Slice" & n - (j - 2) + 1
oData.Filename = STLFilePath & BaseFilename(oPart.fullFilename) & "_" & oBoundaryPatch.Name & ".stl"

Call oSTLTranslator.SaveCopyAs(oPart, oContext, oOptions, oData)
End If

'Convert the boundary patch as surface body
Dim oSurface As WorkSurface
Set oSurface = oBoundaryPatch.SurfaceBody.Parent
oSurface.Visible = False

End If

Next

'Close the completed part file
ThisApplication.Documents.CloseAll

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






















