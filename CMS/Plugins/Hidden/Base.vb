Namespace Plugins.Hidden

    Public MustInherit Class Base
        Implements [Interface]

        Private mobj_page As Page
        Private mstr_plugurl As String
        Private mstr_plugdir As String

        Public Overridable Property Page() As Page Implements [Interface].Page
            Get
                Return mobj_page
            End Get
            Set(ByVal value As Page)
                mobj_page = value
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

    End Class

End Namespace