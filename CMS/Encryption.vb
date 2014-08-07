Imports System.IO
Imports System.Security
Imports System.Security.Principal
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Configuration

'.NET Encryption Simplified

'1. Hashing : Hashes aren't encryption, per se, but they are fundamental to all other encryption operations. A hash is a data fingerprint - a tiny set of bytes that represents the uniqueness of a much larger block of bytes. Like fingerprints, no two should ever be alike, and a matching fingerprint is conclusive proof of identity. A full discussion of hashes is outside the scope of this article, but I highly recommend Steve Friedl's Illustrated Guide to Cryptographic Hashes for more background.
'2. Symmetric Encryption : In symmetric encryption, a single key is used for encrypting and decrypting the data. This type of encryption is quite fast, but has a severe problem: in order to share a secret with someone, they have to know your key. This implies a very high level of trust between people sharing secrets; if an unscrupulous person has your key-- or if your key is intercepted by a spy-- they can decrypt all the messages you send using that key!

'Hash
'Provider   Length (bits)   Security 	Speed
'CRC32 	    32 	            low 	    fast
'SHA1 	    160 	        moderate 	medium
'SHA256 	256 	        high 	    slow
'SHA384 	384 	        high 	    slow
'SHA512 	512 	        extreme 	slow
'MD5 	    128 	        moderate 	medium

'Symmetric
'Provider   Length (bits)   Known Vulnerabilities
'DES        64              yes
'RC2        40-128          yes
'Rijndael   128, 192, 256   no
'TripleDES  128, 192        no

'Encryption.Symmetric - single key
'Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.MD5)
'Dim key As New Encryption.Data("My Password")
'Dim encryptedData As Encryption.Data
'encryptedData = sym.Encrypt(New Encryption.Data("Data to encrypt"), key)
'Dim base64EncryptedString As String = encryptedData.ToBase64

'Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.MD5)
'Dim key As New Encryption.Data("My Password")
'Dim encryptedData As New Encryption.Data
'encryptedData.Base64 = base64EncryptedString
'Dim decryptedData As Encryption.Data
'decryptedData = sym.Decrypt(encryptedData, key)

Namespace Encryption

#Region " Hash "

    Public Class Hash

        Public Enum Provider
            CRC32 ''' Cyclic Redundancy Check provider, 32-bit
            SHA1 ''' Secure Hashing Algorithm provider, SHA-1 variant, 160-bit
            SHA256 ''' Secure Hashing Algorithm provider, SHA-2 variant, 256-bit
            SHA384 ''' Secure Hashing Algorithm provider, SHA-2 variant, 384-bit
            SHA512 ''' Secure Hashing Algorithm provider, SHA-2 variant, 512-bit
            MD5 ''' Message Digest algorithm 5, 128-bit
        End Enum

        Private _Hash As HashAlgorithm
        Private _HashValue As New Data

        Private Sub New()
        End Sub

        Public Sub New(ByVal p As Provider)
            Select Case p
                Case Provider.CRC32
                    _Hash = New CRC32
                Case Provider.MD5
                    _Hash = New MD5CryptoServiceProvider
                Case Provider.SHA1
                    _Hash = New SHA1Managed
                Case Provider.SHA256
                    _Hash = New SHA256Managed
                Case Provider.SHA384
                    _Hash = New SHA384Managed
                Case Provider.SHA512
                    _Hash = New SHA512Managed
            End Select
        End Sub

        Public ReadOnly Property Value() As Data
            Get
                Return _HashValue
            End Get
        End Property

        Public Function Calculate(ByRef s As System.IO.Stream) As Data
            _HashValue.Bytes = _Hash.ComputeHash(s)
            Return _HashValue
        End Function

        Public Function Calculate(ByVal d As Data) As Data
            Return CalculatePrivate(d.Bytes)
        End Function

        Public Function Calculate(ByVal d As Data, ByVal salt As Data) As Data
            Dim nb(d.Bytes.Length + salt.Bytes.Length - 1) As Byte
            salt.Bytes.CopyTo(nb, 0)
            d.Bytes.CopyTo(nb, salt.Bytes.Length)
            Return CalculatePrivate(nb)
        End Function

        Private Function CalculatePrivate(ByVal b() As Byte) As Data
            _HashValue.Bytes = _Hash.ComputeHash(b)
            Return _HashValue
        End Function

#Region "  CRC32 HashAlgorithm"
        Private Class CRC32
            Inherits HashAlgorithm

            Private result As Integer = &HFFFFFFFF

            Protected Overrides Sub HashCore(ByVal array() As Byte, ByVal ibStart As Integer, ByVal cbSize As Integer)
                Dim lookup As Integer
                For i As Integer = ibStart To cbSize - 1
                    lookup = (result And &HFF) Xor array(i)
                    result = ((result And &HFFFFFF00) \ &H100) And &HFFFFFF
                    result = result Xor crcLookup(lookup)
                Next i
            End Sub

            Protected Overrides Function HashFinal() As Byte()
                Dim b() As Byte = BitConverter.GetBytes(Not result)
                Array.Reverse(b)
                Return b
            End Function

            Public Overrides Sub Initialize()
                result = &HFFFFFFFF
            End Sub

            Private crcLookup() As Integer = { _
                &H0, &H77073096, &HEE0E612C, &H990951BA, _
                &H76DC419, &H706AF48F, &HE963A535, &H9E6495A3, _
                &HEDB8832, &H79DCB8A4, &HE0D5E91E, &H97D2D988, _
                &H9B64C2B, &H7EB17CBD, &HE7B82D07, &H90BF1D91, _
                &H1DB71064, &H6AB020F2, &HF3B97148, &H84BE41DE, _
                &H1ADAD47D, &H6DDDE4EB, &HF4D4B551, &H83D385C7, _
                &H136C9856, &H646BA8C0, &HFD62F97A, &H8A65C9EC, _
                &H14015C4F, &H63066CD9, &HFA0F3D63, &H8D080DF5, _
                &H3B6E20C8, &H4C69105E, &HD56041E4, &HA2677172, _
                &H3C03E4D1, &H4B04D447, &HD20D85FD, &HA50AB56B, _
                &H35B5A8FA, &H42B2986C, &HDBBBC9D6, &HACBCF940, _
                &H32D86CE3, &H45DF5C75, &HDCD60DCF, &HABD13D59, _
                &H26D930AC, &H51DE003A, &HC8D75180, &HBFD06116, _
                &H21B4F4B5, &H56B3C423, &HCFBA9599, &HB8BDA50F, _
                &H2802B89E, &H5F058808, &HC60CD9B2, &HB10BE924, _
                &H2F6F7C87, &H58684C11, &HC1611DAB, &HB6662D3D, _
                &H76DC4190, &H1DB7106, &H98D220BC, &HEFD5102A, _
                &H71B18589, &H6B6B51F, &H9FBFE4A5, &HE8B8D433, _
                &H7807C9A2, &HF00F934, &H9609A88E, &HE10E9818, _
                &H7F6A0DBB, &H86D3D2D, &H91646C97, &HE6635C01, _
                &H6B6B51F4, &H1C6C6162, &H856530D8, &HF262004E, _
                &H6C0695ED, &H1B01A57B, &H8208F4C1, &HF50FC457, _
                &H65B0D9C6, &H12B7E950, &H8BBEB8EA, &HFCB9887C, _
                &H62DD1DDF, &H15DA2D49, &H8CD37CF3, &HFBD44C65, _
                &H4DB26158, &H3AB551CE, &HA3BC0074, &HD4BB30E2, _
                &H4ADFA541, &H3DD895D7, &HA4D1C46D, &HD3D6F4FB, _
                &H4369E96A, &H346ED9FC, &HAD678846, &HDA60B8D0, _
                &H44042D73, &H33031DE5, &HAA0A4C5F, &HDD0D7CC9, _
                &H5005713C, &H270241AA, &HBE0B1010, &HC90C2086, _
                &H5768B525, &H206F85B3, &HB966D409, &HCE61E49F, _
                &H5EDEF90E, &H29D9C998, &HB0D09822, &HC7D7A8B4, _
                &H59B33D17, &H2EB40D81, &HB7BD5C3B, &HC0BA6CAD, _
                &HEDB88320, &H9ABFB3B6, &H3B6E20C, &H74B1D29A, _
                &HEAD54739, &H9DD277AF, &H4DB2615, &H73DC1683, _
                &HE3630B12, &H94643B84, &HD6D6A3E, &H7A6A5AA8, _
                &HE40ECF0B, &H9309FF9D, &HA00AE27, &H7D079EB1, _
                &HF00F9344, &H8708A3D2, &H1E01F268, &H6906C2FE, _
                &HF762575D, &H806567CB, &H196C3671, &H6E6B06E7, _
                &HFED41B76, &H89D32BE0, &H10DA7A5A, &H67DD4ACC, _
                &HF9B9DF6F, &H8EBEEFF9, &H17B7BE43, &H60B08ED5, _
                &HD6D6A3E8, &HA1D1937E, &H38D8C2C4, &H4FDFF252, _
                &HD1BB67F1, &HA6BC5767, &H3FB506DD, &H48B2364B, _
                &HD80D2BDA, &HAF0A1B4C, &H36034AF6, &H41047A60, _
                &HDF60EFC3, &HA867DF55, &H316E8EEF, &H4669BE79, _
                &HCB61B38C, &HBC66831A, &H256FD2A0, &H5268E236, _
                &HCC0C7795, &HBB0B4703, &H220216B9, &H5505262F, _
                &HC5BA3BBE, &HB2BD0B28, &H2BB45A92, &H5CB36A04, _
                &HC2D7FFA7, &HB5D0CF31, &H2CD99E8B, &H5BDEAE1D, _
                &H9B64C2B0, &HEC63F226, &H756AA39C, &H26D930A, _
                &H9C0906A9, &HEB0E363F, &H72076785, &H5005713, _
                &H95BF4A82, &HE2B87A14, &H7BB12BAE, &HCB61B38, _
                &H92D28E9B, &HE5D5BE0D, &H7CDCEFB7, &HBDBDF21, _
                &H86D3D2D4, &HF1D4E242, &H68DDB3F8, &H1FDA836E, _
                &H81BE16CD, &HF6B9265B, &H6FB077E1, &H18B74777, _
                &H88085AE6, &HFF0F6A70, &H66063BCA, &H11010B5C, _
                &H8F659EFF, &HF862AE69, &H616BFFD3, &H166CCF45, _
                &HA00AE278, &HD70DD2EE, &H4E048354, &H3903B3C2, _
                &HA7672661, &HD06016F7, &H4969474D, &H3E6E77DB, _
                &HAED16A4A, &HD9D65ADC, &H40DF0B66, &H37D83BF0, _
                &HA9BCAE53, &HDEBB9EC5, &H47B2CF7F, &H30B5FFE9, _
                &HBDBDF21C, &HCABAC28A, &H53B39330, &H24B4A3A6, _
                &HBAD03605, &HCDD70693, &H54DE5729, &H23D967BF, _
                &HB3667A2E, &HC4614AB8, &H5D681B02, &H2A6F2B94, _
                &HB40BBE37, &HC30C8EA1, &H5A05DF1B, &H2D02EF8D}

            Public Overrides ReadOnly Property Hash() As Byte()
                Get
                    Dim b() As Byte = BitConverter.GetBytes(Not result)
                    Array.Reverse(b)
                    Return b
                End Get
            End Property
        End Class

#End Region

    End Class
#End Region

#Region " Symmetric "

    Public Class Symmetric
        Private Const _DefaultIntializationVector As String = "%1Az=-@qT"
        Private Const _BufferSize As Integer = 2048

        Public Enum Provider
            DES
            RC2
            Rijndael
            TripleDES
        End Enum

        Private _data As Data
        Private _key As Data
        Private _iv As Data
        Private _crypto As SymmetricAlgorithm
        Private _EncryptedBytes As Byte()
        Private _UseDefaultInitializationVector As Boolean

        Private Sub New()
        End Sub

        Public Sub New(ByVal provider As Provider, Optional ByVal useDefaultInitializationVector As Boolean = True)
            Select Case provider
                Case provider.DES
                    _crypto = New DESCryptoServiceProvider
                Case provider.RC2
                    _crypto = New RC2CryptoServiceProvider
                Case provider.Rijndael
                    _crypto = New RijndaelManaged
                Case provider.TripleDES
                    _crypto = New TripleDESCryptoServiceProvider
            End Select
            Me.Key = RandomKey()
            If useDefaultInitializationVector Then
                Me.IntializationVector = New Data(_DefaultIntializationVector)
            Else
                Me.IntializationVector = RandomInitializationVector()
            End If
        End Sub

        Public Property KeySizeBytes() As Integer
            Get
                Return _crypto.KeySize \ 8
            End Get
            Set(ByVal Value As Integer)
                _crypto.KeySize = Value * 8
                _key.MaxBytes = Value
            End Set
        End Property

        Public Property KeySizeBits() As Integer
            Get
                Return _crypto.KeySize
            End Get
            Set(ByVal Value As Integer)
                _crypto.KeySize = Value
                _key.MaxBits = Value
            End Set
        End Property

        Public Property Key() As Data
            Get
                Return _key
            End Get
            Set(ByVal Value As Data)
                _key = Value
                _key.MaxBytes = _crypto.LegalKeySizes(0).MaxSize \ 8
                _key.MinBytes = _crypto.LegalKeySizes(0).MinSize \ 8
                _key.StepBytes = _crypto.LegalKeySizes(0).SkipSize \ 8
            End Set
        End Property

        Public Property IntializationVector() As Data
            Get
                Return _iv
            End Get
            Set(ByVal Value As Data)
                _iv = Value
                _iv.MaxBytes = _crypto.BlockSize \ 8
                _iv.MinBytes = _crypto.BlockSize \ 8
            End Set
        End Property

        Public Function RandomInitializationVector() As Data
            _crypto.GenerateIV()
            Dim d As New Data(_crypto.IV)
            Return d
        End Function

        Public Function RandomKey() As Data
            _crypto.GenerateKey()
            Dim d As New Data(_crypto.Key)
            Return d
        End Function

        Private Sub ValidateKeyAndIv(ByVal isEncrypting As Boolean)
            If _key.IsEmpty Then
                If isEncrypting Then
                    _key = RandomKey()
                Else
                    Throw New CryptographicException("No key was provided for the decryption operation!")
                End If
            End If
            If _iv.IsEmpty Then
                If isEncrypting Then
                    _iv = RandomInitializationVector()
                Else
                    Throw New CryptographicException("No initialization vector was provided for the decryption operation!")
                End If
            End If
            _crypto.Key = _key.Bytes
            _crypto.IV = _iv.Bytes
        End Sub

        Public Function Encrypt(ByVal d As Data, ByVal key As Data) As Data
            Me.Key = key
            Return Encrypt(d)
        End Function

        Public Function Encrypt(ByVal d As Data) As Data
            Using ms As New IO.MemoryStream
                ValidateKeyAndIv(True)
                Using cs As New CryptoStream(ms, _crypto.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(d.Bytes, 0, d.Bytes.Length)
                End Using
                Return New Data(ms.ToArray)
            End Using
        End Function

        Public Function Encrypt(ByVal s As Stream, ByVal key As Data, ByVal iv As Data) As Data
            Me.IntializationVector = iv
            Me.Key = key
            Return Encrypt(s)
        End Function

        Public Function Encrypt(ByVal s As Stream, ByVal key As Data) As Data
            Me.Key = key
            Return Encrypt(s)
        End Function

        Public Function Encrypt(ByVal s As Stream) As Data
            Using ms As New IO.MemoryStream
                Dim b(_BufferSize) As Byte
                ValidateKeyAndIv(True)
                Using cs As New CryptoStream(ms, _crypto.CreateEncryptor(), CryptoStreamMode.Write)
                    Dim i As Integer = s.Read(b, 0, _BufferSize)
                    Do While i > 0
                        cs.Write(b, 0, i)
                        i = s.Read(b, 0, _BufferSize)
                    Loop
                End Using
                Return New Data(ms.ToArray)
            End Using
        End Function

        Public Function Decrypt(ByVal encryptedData As Data, ByVal key As Data) As Data
            Me.Key = key
            Return Decrypt(encryptedData)
        End Function

        Public Function Decrypt(ByVal encryptedStream As Stream, ByVal key As Data) As Data
            Me.Key = key
            Return Decrypt(encryptedStream)
        End Function

        Public Function Decrypt(ByVal encryptedStream As Stream) As Data
            Using ms As New System.IO.MemoryStream
                Dim b(_BufferSize) As Byte
                ValidateKeyAndIv(False)
                Using cs As New CryptoStream(encryptedStream, _crypto.CreateDecryptor(), CryptoStreamMode.Read)
                    Dim i As Integer = cs.Read(b, 0, _BufferSize)
                    Do While i > 0
                        ms.Write(b, 0, i)
                        i = cs.Read(b, 0, _BufferSize)
                    Loop
                End Using
                Return New Data(ms.ToArray)
            End Using
        End Function

        Public Function Decrypt(ByVal encryptedData As Data) As Data
            Using ms As New System.IO.MemoryStream(encryptedData.Bytes, 0, encryptedData.Bytes.Length)
                Dim b() As Byte = New Byte(encryptedData.Bytes.Length - 1) {}
                ValidateKeyAndIv(False)
                Using cs As New CryptoStream(ms, _crypto.CreateDecryptor(), CryptoStreamMode.Read)
                    Try
                        cs.Read(b, 0, encryptedData.Bytes.Length - 1)
                    Catch ex As CryptographicException
                        Throw New CryptographicException("Unable to decrypt data. The provided key may be invalid.", ex)
                    End Try
                End Using
                Return New Data(b)
            End Using
        End Function

    End Class

#End Region

#Region " Data "

    Public Class Data
        Private _b As Byte()
        Private _MaxBytes As Integer = 0
        Private _MinBytes As Integer = 0
        Private _StepBytes As Integer = 0

        Public Shared DefaultEncoding As Text.Encoding = System.Text.Encoding.GetEncoding("Windows-1252")
        Public Encoding As Text.Encoding = DefaultEncoding

        Public Sub New()
        End Sub

        Public Sub New(ByVal b As Byte())
            _b = b
        End Sub

        Public Sub New(ByVal s As String)
            Me.Text = s
        End Sub

        Public Sub New(ByVal s As String, ByVal encoding As System.Text.Encoding)
            Me.Encoding = encoding
            Me.Text = s
        End Sub

        Public ReadOnly Property IsEmpty() As Boolean
            Get
                Return (_b Is Nothing Or _b.Length = 0)
            End Get
        End Property

        Public Property StepBytes() As Integer
            Get
                Return _StepBytes
            End Get
            Set(ByVal Value As Integer)
                _StepBytes = Value
            End Set
        End Property

        Public Property StepBits() As Integer
            Get
                Return _StepBytes * 8
            End Get
            Set(ByVal Value As Integer)
                _StepBytes = Value \ 8
            End Set
        End Property


        Public Property MinBytes() As Integer
            Get
                Return _MinBytes
            End Get
            Set(ByVal Value As Integer)
                _MinBytes = Value
            End Set
        End Property

        Public Property MinBits() As Integer
            Get
                Return _MinBytes * 8
            End Get
            Set(ByVal Value As Integer)
                _MinBytes = Value \ 8
            End Set
        End Property

        Public Property MaxBytes() As Integer
            Get
                Return _MaxBytes
            End Get
            Set(ByVal Value As Integer)
                _MaxBytes = Value
            End Set
        End Property

        Public Property MaxBits() As Integer
            Get
                Return _MaxBytes * 8
            End Get
            Set(ByVal Value As Integer)
                _MaxBytes = Value \ 8
            End Set
        End Property

        Public Property Bytes() As Byte()
            Get
                If _MaxBytes > 0 Then
                    If _b.Length > _MaxBytes Then
                        Dim b(_MaxBytes - 1) As Byte
                        Array.Copy(_b, b, b.Length)
                        _b = b
                    End If
                End If
                If _MinBytes > 0 Then
                    If _b.Length < _MinBytes Then
                        Dim b(_MinBytes - 1) As Byte
                        Array.Copy(_b, b, _b.Length)
                        _b = b
                    End If
                End If
                Return _b
            End Get
            Set(ByVal Value As Byte())
                _b = Value
            End Set
        End Property

        Public Property Text() As String
            Get
                If _b Is Nothing Then
                    Return ""
                Else
                    Dim i As Integer = Array.IndexOf(_b, CType(0, Byte))
                    If i >= 0 Then
                        Return Me.Encoding.GetString(_b, 0, i)
                    Else
                        Return Me.Encoding.GetString(_b)
                    End If
                End If
            End Get
            Set(ByVal Value As String)
                _b = Me.Encoding.GetBytes(Value)
            End Set
        End Property

        Public Property Hex() As String
            Get
                Return Utils.ToHex(_b)
            End Get
            Set(ByVal Value As String)
                _b = Utils.FromHex(Value)
            End Set
        End Property

        Public Property Base64() As String
            Get
                Return Utils.ToBase64(_b)
            End Get
            Set(ByVal Value As String)
                _b = Utils.FromBase64(Value)
            End Set
        End Property

        Public Shadows Function ToString() As String
            Return Me.Text
        End Function

        Public Function ToBase64() As String
            Return Me.Base64
        End Function

        Public Function ToHex() As String
            Return Me.Hex
        End Function

    End Class

#End Region

#Region " Utils "

    Friend Class Utils

        Friend Shared Function ToHex(ByVal ba() As Byte) As String
            If ba Is Nothing OrElse ba.Length = 0 Then
                Return ""
            End If
            Const HexFormat As String = "{0:X2}"
            Dim sb As New StringBuilder
            For Each b As Byte In ba
                sb.Append(String.Format(HexFormat, b))
            Next
            Return sb.ToString
        End Function

        Friend Shared Function FromHex(ByVal hexEncoded As String) As Byte()
            If hexEncoded Is Nothing OrElse hexEncoded.Length = 0 Then
                Return Nothing
            End If
            Try
                Dim l As Integer = Convert.ToInt32(hexEncoded.Length / 2)
                Dim b(l - 1) As Byte
                For i As Integer = 0 To l - 1
                    b(i) = Convert.ToByte(hexEncoded.Substring(i * 2, 2), 16)
                Next
                Return b
            Catch ex As Exception
                Throw New System.FormatException("The provided string does not appear to be Hex encoded:" & _
                    Environment.NewLine & hexEncoded & Environment.NewLine, ex)
            End Try
        End Function

        Friend Shared Function FromBase64(ByVal base64Encoded As String) As Byte()
            If base64Encoded Is Nothing OrElse base64Encoded.Length = 0 Then
                Return Nothing
            End If
            Try
                Return Convert.FromBase64String(base64Encoded)
            Catch ex As System.FormatException
                Throw New System.FormatException("The provided string does not appear to be Base64 encoded:" & _
                    Environment.NewLine & base64Encoded & Environment.NewLine, ex)
            End Try
        End Function

        Friend Shared Function ToBase64(ByVal b() As Byte) As String
            If b Is Nothing OrElse b.Length = 0 Then
                Return ""
            End If
            Return Convert.ToBase64String(b)
        End Function

    End Class

#End Region

End Namespace