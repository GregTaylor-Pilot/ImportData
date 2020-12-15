Imports System
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography


Public Class SecureIt

    Private MyInitialVector As String
    Private MyDesiredEncryption As String
    Private MyIterations As Integer

    Public Enum MyEncryptionTypes
        SHA1
        MD5
    End Enum


    Public Sub New(ByVal Vector As String, Optional ByVal MyEncryptionType As SecureIt.MyEncryptionTypes = SecureIt.MyEncryptionTypes.SHA1,
                   Optional ByVal Iterations As Integer = 2)

        Dim VectorLen As Integer = Vector.Length
        If VectorLen > 16 Then
            Vector = Vector.Substring(0, 16)
        ElseIf VectorLen < 16 Then
            Vector = Vector.PadLeft(16, "P"c)
        Else
            'Do Nothing
        End If

        MyInitialVector = Vector
        MyIterations = Iterations

        Select Case MyEncryptionType
            Case SecureIt.MyEncryptionTypes.SHA1
                MyDesiredEncryption = "SHA1"
            Case SecureIt.MyEncryptionTypes.MD5
                MyDesiredEncryption = "MD5"
            Case Else
                MyDesiredEncryption = "SHA1"

        End Select


    End Sub

    Public Function AESEncrypt(ByVal PlainText As String, ByVal Password As String, ByVal salt As String) As String
        Dim HashAlgorithm As String = MyDesiredEncryption 'Can be SHA1 or MD5
        Dim PasswordIterations As Integer = MyIterations
        Dim InitialVector As String = MyInitialVector 'This should be a string of 16 ASCII characters.
        Dim KeySize As Integer = 256 'Can be 128, 192, or 256.

        If (String.IsNullOrEmpty(PlainText)) Then
            Return ""
            Exit Function
        End If
        Dim InitialVectorBytes As Byte() = Encoding.ASCII.GetBytes(InitialVector)
        Dim SaltValueBytes As Byte() = Encoding.ASCII.GetBytes(salt)
        Dim PlainTextBytes As Byte() = Encoding.UTF8.GetBytes(PlainText)
        Dim DerivedPassword As DeriveBytes = New PasswordDeriveBytes(Password, SaltValueBytes, HashAlgorithm, PasswordIterations)
        Dim KeyBytes As Byte() = DerivedPassword.GetBytes(CInt(KeySize / 8))
        Dim SymmetricKey As RijndaelManaged = New RijndaelManaged()
        SymmetricKey.Mode = CipherMode.CBC

        Dim CipherTextBytes As Byte() = Nothing
        Using Encryptor As ICryptoTransform = SymmetricKey.CreateEncryptor(KeyBytes, InitialVectorBytes)
            Using MemStream As New MemoryStream()
                Using CryptoStream As New CryptoStream(MemStream, Encryptor, CryptoStreamMode.Write)
                    CryptoStream.Write(PlainTextBytes, 0, PlainTextBytes.Length)
                    CryptoStream.FlushFinalBlock()
                    CipherTextBytes = MemStream.ToArray()
                    MemStream.Close()
                    CryptoStream.Close()
                End Using
            End Using
        End Using
        SymmetricKey.Clear()
        Return Convert.ToBase64String(CipherTextBytes)
    End Function

    Public Function AESDecrypt(ByVal CipherText As String, ByVal password As String, ByVal salt As String) As String
        Dim HashAlgorithm As String = MyDesiredEncryption
        Dim PasswordIterations As Integer = MyIterations
        Dim InitialVector As String = MyInitialVector
        Dim KeySize As Integer = 256

        If (String.IsNullOrEmpty(CipherText)) Then
            Return ""
        End If
        Dim InitialVectorBytes As Byte() = Encoding.ASCII.GetBytes(InitialVector)
        Dim SaltValueBytes As Byte() = Encoding.ASCII.GetBytes(salt)
        Dim CipherTextBytes As Byte() = Convert.FromBase64String(CipherText)
        Dim DerivedPassword As DeriveBytes = New PasswordDeriveBytes(password, SaltValueBytes, HashAlgorithm, PasswordIterations)
        Dim KeyBytes As Byte() = DerivedPassword.GetBytes(CInt(KeySize / 8))
        Dim SymmetricKey As RijndaelManaged = New RijndaelManaged()
        SymmetricKey.Mode = CipherMode.CBC
        Dim PlainTextBytes As Byte() = New Byte(CipherTextBytes.Length - 1) {}

        Dim ByteCount As Integer = 0

        Using Decryptor As ICryptoTransform = SymmetricKey.CreateDecryptor(KeyBytes, InitialVectorBytes)
            Using MemStream As MemoryStream = New MemoryStream(CipherTextBytes)
                Using CryptoStream As CryptoStream = New CryptoStream(MemStream, Decryptor, CryptoStreamMode.Read)
                    ByteCount = CryptoStream.Read(PlainTextBytes, 0, PlainTextBytes.Length)
                    MemStream.Close()
                    CryptoStream.Close()
                End Using
            End Using
        End Using
        SymmetricKey.Clear()
        Return Encoding.UTF8.GetString(PlainTextBytes, 0, ByteCount)
    End Function

End Class
