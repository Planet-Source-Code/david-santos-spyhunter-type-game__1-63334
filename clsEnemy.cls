VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'local variable(s) to hold property value(s)
Private mvarx As Single 'local copy
Private mvary As Single 'local copy
Private mvardx As Single 'local copy
Private mvardy As Single 'local copy
Private mvartype As Long 'local copy
'local variable(s) to hold property value(s)
Private mvardead As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarTag As String 'local copy
'local variable(s) to hold property value(s)
Private mvarexpframe As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarprevTile As Byte 'local copy
Public Property Let prevTile(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.prevTile = 5
    mvarprevTile = vData
End Property


Public Property Get prevTile() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.prevTile
    prevTile = mvarprevTile
End Property



Public Property Let expframe(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.expframe = 5
    mvarexpframe = vData
End Property


Public Property Set expframe(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.expframe = Form1
    Set mvarexpframe = vData
End Property


Public Property Get expframe() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.expframe
    If IsObject(mvarexpframe) Then
        Set expframe = mvarexpframe
    Else
        expframe = mvarexpframe
    End If
End Property



Public Property Let Tag(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tag = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
End Property



Public Property Let dead(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dead = 5
    mvardead = vData
End Property


Public Property Get dead() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dead
    dead = mvardead
End Property




Public Property Let mtype(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.type = 5
    mvartype = vData
End Property


Public Property Get mtype() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.type
    mtype = mvartype
End Property



Public Property Let dy(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dy = 5
    mvardy = vData
End Property


Public Property Get dy() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dy
    dy = mvardy
End Property



Public Property Let dx(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dx = 5
    mvardx = vData
End Property


Public Property Get dx() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dx
    dx = mvardx
End Property



Public Property Let y(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.y = 5
    mvary = vData
End Property


Public Property Get y() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.y
    y = mvary
End Property



Public Property Let x(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.x = 5
    mvarx = vData
End Property


Public Property Get x() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.x
    x = mvarx
End Property



