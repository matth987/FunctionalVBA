VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FunctionObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_funcName As String
Private m_owner As Object
Private m_callType As VbCallType
Private m_boundArgs() As Variant

Private Sub Class_Initialize()
    m_boundArgs = MakeEmptyParamArray
End Sub

Private Function MakeEmptyParamArray(ParamArray dummy() As Variant) As Variant
    MakeEmptyParamArray = dummy
End Function

Private Sub MakeThisThat(ByRef this As Variant, that As Variant)
    If IsObject(that) Then
        Set this = that
    Else
        this = that
    End If
End Sub

Public Property Get FunctionName() As String
    FunctionName = m_funcName
End Property

Public Property Get owner() As Object
    Set owner = m_owner
End Property

Public Sub WrapFunction(funcName As String, Optional owner As Object = Nothing, Optional callType As VbCallType = VbMethod)
    m_funcName = funcName
    Set m_owner = owner
    m_callType = callType
End Sub

Public Sub BindArgs(ParamArray args() As Variant)
    m_boundArgs = args
End Sub

Private Function PrepArgs(ByRef moreArgs As Variant) As Variant
    Dim boundSize As Integer
    boundSize = UBound(m_boundArgs) - LBound(m_boundArgs) + 1
    
    Dim moreSize As Integer
    moreSize = UBound(moreArgs) - LBound(moreArgs) + 1
    
    Dim compiledSize As Integer
    compiledSize = boundSize + moreSize
    
    If compiledSize < 1 Then
        PrepArgs = MakeEmptyParamArray
        Exit Function
    End If
    
    If compiledSize > 30 Then _
        Err.Raise vbObjectError, , "Only up to 30 arguments may be provided."

    If boundSize < 1 Or moreSize < 1 Then
        PrepArgs = IIf(boundSize > 0, m_boundArgs, moreArgs)
        Exit Function
    End If
    
    Dim compiledArgs() As Variant
    ReDim compiledArgs(0 To compiledSize - 1)
    
    Dim i As Integer
    For i = 0 To boundSize - 1
        Call MakeThisThat(compiledArgs(i), m_boundArgs(i))
    Next
    
    Dim j As Integer
    For j = 0 To moreSize - 1
        Call MakeThisThat(compiledArgs(i), moreArgs(j))
        i = i + 1
    Next
    
    PrepArgs = compiledArgs
    
End Function

Public Function CallFunction(ParamArray args() As Variant) As Variant
    Dim passArgs() As Variant
    passArgs = args
    
    passArgs = PrepArgs(passArgs)

    If m_owner Is Nothing Then
        Call MakeThisThat(CallFunction, CallAppRun(passArgs))
    Else
        Call MakeThisThat(CallFunction, CallMember(passArgs))
    End If

End Function

Private Function CallAppRun(ByRef args() As Variant) As Variant
    Select Case UBound(args)
        Case -1:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName))
        Case 0:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0)))
        Case 1:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1)))
        Case 2:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2)))
        Case 3:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3)))
        Case 4:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4)))
        Case 5:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5)))
        Case 6:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6)))
        Case 7:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7)))
        Case 8:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8)))
        Case 9:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9)))
        Case 10:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10)))
        Case 11:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11)))
        Case 12:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12)))
        Case 13:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13)))
        Case 14:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14)))
        Case 15:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15)))
        Case 16:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16)))
        Case 17:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17)))
        Case 18:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18)))
        Case 19:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19)))
        Case 20:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20)))
        Case 21:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21)))
        Case 22:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22)))
        Case 23:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23)))
        Case 24:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24)))
        Case 25:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25)))
        Case 26:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26)))
        Case 27:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27)))
        Case 28:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28)))
        Case 29:
            Call MakeThisThat(CallAppRun, Application.Run(m_funcName, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28), args(29)))
    End Select
End Function

Private Function CallMember(ByRef args() As Variant) As Variant
    Select Case UBound(args)
        Case -1:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType))
        Case 0:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0)))
        Case 1:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1)))
        Case 2:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2)))
        Case 3:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3)))
        Case 4:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4)))
        Case 5:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5)))
        Case 6:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6)))
        Case 7:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7)))
        Case 8:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8)))
        Case 9:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9)))
        Case 10:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10)))
        Case 11:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11)))
        Case 12:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12)))
        Case 13:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13)))
        Case 14:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14)))
        Case 15:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15)))
        Case 16:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16)))
        Case 17:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17)))
        Case 18:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18)))
        Case 19:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19)))
        Case 20:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20)))
        Case 21:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21)))
        Case 22:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22)))
        Case 23:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23)))
        Case 24:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24)))
        Case 25:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25)))
        Case 26:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26)))
        Case 27:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27)))
        Case 28:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28)))
        Case 29:
            Call MakeThisThat(CallMember, CallByName(m_owner, m_funcName, m_callType, args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12), args(13), args(14), args(15), args(16), args(17), args(18), args(19), args(20), args(21), args(22), args(23), args(24), args(25), args(26), args(27), args(28), args(29)))
    End Select
End Function
