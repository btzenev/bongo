Namespace Plugins.Client

    Public MustInherit Class Base
        Implements [Interface]

        Private mobj_page As Page
        Private mint_mod As Integer
        Private mstr_plugurl As String
        Private mstr_plugdir As String

        Public MustOverride ReadOnly Property Name() As String Implements [Interface].Name

        Public Overridable ReadOnly Property Visible() As Boolean Implements [Interface].Visible
            Get
                Return True
            End Get
        End Property

        Public Overridable ReadOnly Property Head() As String Implements [Interface].Head
            Get
                Return String.Empty
            End Get
        End Property

        Public Overridable Property Page() As Page Implements [Interface].Page
            Get
                Return mobj_page
            End Get
            Set(ByVal value As Page)
                mobj_page = value
            End Set
        End Property

        Public Overridable Property ModuleID() As Integer Implements [Interface].ModuleID
            Get
                Return mint_mod
            End Get
            Set(ByVal value As Integer)
                mint_mod = value
            End Set
        End Property

        Public Overridable Property PluginDir() As String Implements [Interface].PluginDir
            Get
                Return mstr_plugdir
            End Get
            Set(ByVal value As String)
                mstr_plugdir = value
            End Set
        End Property

        Public Overridable Property PluginUrl() As String Implements [Interface].PluginUrl
            Get
                Return mstr_plugurl
            End Get
            Set(ByVal value As String)
                mstr_plugurl = value
            End Set
        End Property

        Public MustOverride Function Render() As String Implements [Interface].Render

        Public Overridable Function Add() As Boolean Implements [Interface].Add
            'nothing
        End Function

        Public Overridable Function Delete() As Boolean Implements [Interface].Delete
            'nothing
        End Function

        Public Overridable Function Edit() As String Implements [Interface].Edit
            Return "This module does not require any additional information to operate."
        End Function

    End Class

End Namespace