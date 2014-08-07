Imports System.IO
Imports System.Text
Imports System.Web.Mail
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Configuration

Public Class Cache

    Public Shared ReadOnly Property [Get](ByVal vstr_key As String) As Object
        Get
            On Error Resume Next
            If Debug Then Return Nothing
            Return HttpContext.Current.Cache(vstr_key)
        End Get
  End Property

  Public Shared Sub Add(ByVal vstr_key As String, ByVal vobj_data As Object)
    Add(vstr_key, vobj_data, TimeSpan.FromHours(1))
  End Sub

  Public Shared Sub Add(ByVal vstr_key As String, ByVal vobj_data As Object, ByVal vobj_span As TimeSpan)
    On Error Resume Next
    Remove(vstr_key)
    HttpContext.Current.Cache.Insert(vstr_key, vobj_data, Nothing, Web.Caching.Cache.NoAbsoluteExpiration, vobj_span)
  End Sub

    Public Shared Sub Remove(ByVal vstr_key As String)
        On Error Resume Next
        HttpContext.Current.Cache.Remove(vstr_key)
    End Sub

    Public Shared Sub Insert(ByVal vstr_key As String, ByVal value As Object, ByVal dependencies As System.Web.Caching.CacheDependency)
        On Error Resume Next
        Remove(vstr_key)
        HttpContext.Current.Cache.Insert(vstr_key, value, dependencies)
    End Sub

    Public Shared Sub Insert(ByVal vstr_key As String, ByVal value As Object, ByVal dependencies As System.Web.Caching.CacheDependency, ByVal absoluteExpiration As Date, ByVal slidingExpiration As System.TimeSpan)
        On Error Resume Next
        Remove(vstr_key)
        HttpContext.Current.Cache.Insert(vstr_key, value, dependencies, absoluteExpiration, slidingExpiration)
    End Sub

    Public Shared Sub Insert(ByVal vstr_key As String, ByVal value As Object, ByVal dependencies As System.Web.Caching.CacheDependency, ByVal absoluteExpiration As Date, ByVal slidingExpiration As System.TimeSpan, ByVal priority As System.Web.Caching.CacheItemPriority, ByVal onRemoveCallback As System.Web.Caching.CacheItemRemovedCallback)
        On Error Resume Next
        Remove(vstr_key)
        HttpContext.Current.Cache.Insert(vstr_key, value, dependencies, absoluteExpiration, slidingExpiration, priority, onRemoveCallback)
    End Sub

    Public Shared Sub Insert(ByVal vstr_key As String, ByVal value As Object, ByVal dependencies As System.Web.Caching.CacheDependency, ByVal absoluteExpiration As Date, ByVal slidingExpiration As System.TimeSpan, ByVal onUpdateCallback As System.Web.Caching.CacheItemUpdateCallback)
        On Error Resume Next
        Remove(vstr_key)
        HttpContext.Current.Cache.Insert(vstr_key, value, dependencies, absoluteExpiration, slidingExpiration, onUpdateCallback)
    End Sub

End Class
